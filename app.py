import io
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# ── Constantes ─────────────────────────────────────────────────────────────────
C_HEADER_BG   = "1F4E79"
C_HEADER_FG   = "FFFFFF"
C_CORRECT     = "C6EFCE";  C_CORRECT_FG  = "276221"
C_ERROR       = "FFCCCC";  C_ERROR_FG    = "9C0006"
C_WARN        = "FFEB9C";  C_WARN_FG     = "9C6500"
C_ALT         = "F2F2F2"
C_ONLY_PROP   = "DAE3F3";  C_ONLY_PROP_FG  = "1F4E79"
C_ONLY_CONT   = "FCE4D6";  C_ONLY_CONT_FG  = "843C0C"

KEYWORDS = {
    '0':  ['saldo inicial'],
    '1':  ['canon', 'arrendamiento'],
    '2':  ['admon', 'administra', 'cuota'],
    '3':  ['iva'],
    '4':  ['retefte', 'fuente', 'reten'],
    '5':  ['ica'],
    '6':  ['reteiva', 'rete iva', 'iva'],
    '7':  ['epm', 'servicio', 'public', 'agua', 'tasa'],
    '8':  ['consignac'],
    '9':  ['iva', 'comis'],
    '10': ['4*1000', 'emergencia', 'gmf'],
    '11': ['seguro', 'poliza'],
    '12': ['cree', 'canon'],
    '14': ['iva', 'comis'],
    '15': ['iva', 'ph', 'admon'],
    '16': ['giro'],
    '18': ['derecho', 'contrato'],
    '20': ['comis'],
    '21': ['comis', 'admon'],
    '23': ['fuente', 'comi'],
    '24': ['ica', 'comis'],
    '25': ['iva', 'comis'],
    '26': ['cree', 'comis'],
    '28': ['bomberil', 'bomb', 'impuesto'],
    '29': ['aviso', 'impuesto'],
    '30': ['reparac', 'manteni', 'mtto', 'localiz'],
    '31': ['publicidad'],
    '32': ['predial'],
    '33': ['aseo', 'limpieza', 'anticipo'],
    '34': ['ingreso', 'canon'],
    '35': ['admon', 'ingreso'],
    '36': ['iva', 'ingreso'],
    '37': ['ajuste', 'reconoc', 'indemniz', 'otro'],
    '38': ['descont', 'ingreso', 'giro'],
    '39': ['multa', 'recargo'],
    '40': ['prorrateo', 'gasto'],
    '41': ['bomberil', 'comis'],
    '42': ['aviso', 'comis'],
}

# ── Helpers Excel ──────────────────────────────────────────────────────────────
def _hdr(ws, row, ncols, bg=C_HEADER_BG, fg=C_HEADER_FG):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.font = Font(bold=True, color=fg, name="Arial", size=10)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="thin", color="000000"),
                             right=Side(style="thin", color="CCCCCC"))

def _row(ws, row, ncols, bg=None):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.font = Font(name="Arial", size=9)
        cell.alignment = Alignment(vertical="center")
        if bg:
            cell.fill = PatternFill("solid", fgColor=bg)
        cell.border = Border(bottom=Side(style="thin", color="E0E0E0"),
                             right=Side(style="thin", color="E0E0E0"))

def _color(cell, bg, fg, bold=True):
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.font = Font(name="Arial", size=9, bold=bold, color=fg)

def _widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ── Lógica de conciliación ─────────────────────────────────────────────────────
def run_conciliacion(f_prop, f_cont, f_tc):
    # --- Propietarios activos ---
    df_p = pd.read_excel(f_prop, header=None)
    prop = df_p.iloc[1:]
    actv = prop[prop[12] == 'False'][[2, 13, 14]].copy()
    actv.columns = ["inm", "nombre_prop", "cedula_prop"]
    actv["inm"]         = actv["inm"].astype(str).str.strip()
    actv["nombre_prop"] = actv["nombre_prop"].astype(str).str.strip().str.upper()
    actv["cedula_prop"] = actv["cedula_prop"].astype(str).str.strip()
    prop_inm = actv.drop_duplicates("inm").set_index("inm")

    # --- Contabilidad ---
    df_c = pd.read_excel(f_cont, header=None)
    cont = df_c.iloc[1:].copy()
    cont.columns = range(len(cont.columns))
    cont = cont.rename(columns={3:"nit", 4:"nombre_cont", 6:"detalle",
                                7:"debito", 8:"credito", 12:"inm", 13:"tipo_causa"})
    for col in ["nit", "nombre_cont", "detalle", "inm", "tipo_causa"]:
        cont[col] = cont[col].astype(str).str.strip()
    cont["nombre_cont"] = cont["nombre_cont"].str.upper()
    cont["debito"]  = pd.to_numeric(cont["debito"],  errors="coerce").fillna(0).astype(int)
    cont["credito"] = pd.to_numeric(cont["credito"], errors="coerce").fillna(0).astype(int)
    cont = cont.reset_index(drop=True)

    cont_inm_owners = (
        cont[["inm","nit","nombre_cont"]]
        .drop_duplicates()
        .groupby("inm", group_keys=False)
        .apply(lambda df: df[["nit","nombre_cont"]].values.tolist())
        .to_dict()
    )

    # --- Tipo Causa ---
    df_t = pd.read_excel(f_tc, header=None)
    tc = df_t.iloc[1:].copy()
    tc.columns = ["codigo", "concepto"]
    tc["codigo"]  = tc["codigo"].astype(str).str.strip()
    tc["concepto"]= tc["concepto"].astype(str).str.strip()
    tc_lookup = dict(zip(tc["codigo"], tc["concepto"]))
    tc_lookup["0"] = "SALDO INICIAL"

    # ── Verificación 1 ──────────────────────────────────────────────────────────
    all_inm = sorted(
        set(prop_inm.index) | set(cont_inm_owners.keys()),
        key=lambda x: int(x) if x.isdigit() else float('inf')
    )
    v1_rows = []
    for inm in all_inm:
        in_p = inm in prop_inm.index
        in_c = inm in cont_inm_owners
        n_p  = prop_inm.loc[inm, "nombre_prop"] if in_p else "-"
        ced  = prop_inm.loc[inm, "cedula_prop"] if in_p else "-"

        if not in_c:
            v1_rows.append(dict(
                Inmueble=inm, Propietario_Prop=n_p, NIT_Prop=ced,
                Propietario_Cont="-", NIT_Cont="-",
                Estado="SOLO EN PROPIETARIOS",
                Alerta="Inmueble activo sin movimientos en Contabilidad"))
            continue

        owners = cont_inm_owners[inm]
        if not in_p:
            for nit_c, nom_c in owners:
                v1_rows.append(dict(
                    Inmueble=inm, Propietario_Prop="-", NIT_Prop="-",
                    Propietario_Cont=nom_c, NIT_Cont=nit_c,
                    Estado="SOLO EN CONTABILIDAD",
                    Alerta="Inmueble en Contabilidad sin propietario activo"))
            continue

        cont_nits = [r[0] for r in owners]
        if ced in cont_nits:
            nom_c = next(r[1] for r in owners if r[0] == ced)
            if " ".join(n_p.split()) == " ".join(nom_c.split()):
                estado, alerta = "CORRECTO", ""
            else:
                estado = "ALERTA — NOMBRE DIFIERE"
                alerta = f"NIT coincide pero nombre difiere. Cont: «{nom_c}»"
            v1_rows.append(dict(
                Inmueble=inm, Propietario_Prop=n_p, NIT_Prop=ced,
                Propietario_Cont=nom_c, NIT_Cont=ced,
                Estado=estado, Alerta=alerta))
            for nit_c, nom_c2 in owners:
                if nit_c != ced:
                    v1_rows.append(dict(
                        Inmueble=inm, Propietario_Prop="", NIT_Prop="",
                        Propietario_Cont=nom_c2, NIT_Cont=nit_c,
                        Estado="PROPIETARIO ADICIONAL",
                        Alerta="Contabilidad registra otro propietario en este inmueble"))
        else:
            cont_str = "; ".join(f"{n} ({ni})" for ni, n in owners)
            v1_rows.append(dict(
                Inmueble=inm, Propietario_Prop=n_p, NIT_Prop=ced,
                Propietario_Cont=owners[0][1], NIT_Cont=owners[0][0],
                Estado="ALERTA — NIT NO COINCIDE",
                Alerta=f"NIT {ced} no encontrado. Contabilidad tiene: {cont_str}"))

    df_v1 = pd.DataFrame(v1_rows)
    df_v1["Inmueble"] = pd.to_numeric(df_v1["Inmueble"], errors="coerce")
    df_v1 = df_v1.sort_values("Inmueble").reset_index(drop=True)

    # ── Verificación 2 ──────────────────────────────────────────────────────────
    def check_tc(tc_str, det_str):
        tc_s = str(tc_str).strip()
        det  = det_str.lower()
        conc = tc_lookup.get(tc_s)
        if not conc:
            return "CÓDIGO NO EXISTE", "-", f"TipoCausa '{tc_s}' no está en el catálogo"
        kws = KEYWORDS.get(tc_s, [])
        if not kws:
            return "SIN KEYWORDS", conc, "Sin palabras clave configuradas"
        if any(k in det for k in kws):
            return "CORRECTO", conc, ""
        return "REVISAR", conc, f"Keywords esperadas ({', '.join(kws)}) no encontradas"

    v2_rows = []
    for idx, row in cont.iterrows():
        est, conc, obs = check_tc(row["tipo_causa"], row["detalle"])
        v2_rows.append(dict(
            Fila=idx+2, Inmueble=row["inm"], NIT=row["nit"], Nombre=row["nombre_cont"],
            TipoCausa=row["tipo_causa"], Concepto=conc,
            Detalle=row["detalle"], Estado=est, Observacion=obs))

    df_v2 = pd.DataFrame(v2_rows)
    df_v2["Inmueble"] = pd.to_numeric(df_v2["Inmueble"], errors="coerce")
    df_v2 = df_v2.sort_values(["Inmueble","Fila"]).reset_index(drop=True)

    # ── Verificación 3 — IVA y Retenciones ─────────────────────────────────────
    df_v3 = cont[cont["tipo_causa"].isin(["3","4"])].copy()
    df_v3["inm_num"] = pd.to_numeric(df_v3["inm"], errors="coerce")
    df_v3 = df_v3.sort_values(["inm_num","tipo_causa"]).reset_index(drop=True)
    df_v3["concepto"] = df_v3["tipo_causa"].map(tc_lookup).fillna("-")
    df_v3["neto"] = df_v3["credito"] - df_v3["debito"]

    # ── Verificaciones cruzadas IVA y Retefuente ─────────────────────────────
    def neto_tc(tc_str):
        sub = cont[cont["tipo_causa"] == tc_str]
        return int(sub["credito"].sum() - sub["debito"].sum())

    neto_20    = neto_tc("20")
    neto_14    = neto_tc("14")
    neto_23    = neto_tc("23")
    iva_calc   = round(neto_20 * 0.19)
    rete_calc  = round(neto_20 * 0.11)
    dif_iva    = iva_calc  - neto_14
    dif_rete   = rete_calc - neto_23
    crosscheck = dict(
        neto_20=neto_20, neto_14=neto_14, neto_23=neto_23,
        iva_calc=iva_calc, rete_calc=rete_calc,
        dif_iva=dif_iva, dif_rete=dif_rete,
        ok_iva=(dif_iva==0), ok_rete=(dif_rete==0),
    )

    # ── Análisis de diferencias por inmueble ─────────────────────────────────
    _piv = cont[cont["tipo_causa"].isin(["14","20","23"])].groupby(["inm","tipo_causa"]).apply(
        lambda x: int(x["credito"].sum() - x["debito"].sum()), include_groups=False
    ).unstack(fill_value=0)
    _piv.columns.name = None
    for _c in ["14","20","23"]:
        if _c not in _piv.columns: _piv[_c] = 0
    _piv = _piv[["14","20","23"]].copy()
    _piv.columns = ["neto_14","neto_20","neto_23"]
    _piv["iva_calc"]  = (_piv["neto_20"] * 0.19).round().astype(int)
    _piv["rete_calc"] = (_piv["neto_20"] * 0.11).round().astype(int)
    _piv["dif_iva"]   = _piv["iva_calc"]  - _piv["neto_14"]
    _piv["dif_rete"]  = _piv["rete_calc"] - _piv["neto_23"]
    _piv["inm_num"]   = pd.to_numeric(_piv.index, errors="coerce")
    _nom = cont[cont["tipo_causa"]=="20"].groupby("inm")["nombre_cont"].first()
    _piv["nombre"] = _piv.index.map(_nom)

    def _clr(row):
        if row["dif_rete"] == 0:                return "OK"
        if row["neto_23"] == 0:                 return "SIN TC23"
        if row["neto_23"] == -row["rete_calc"]: return "SIGNO INVERTIDO"
        return "DIFERENCIA PARCIAL"

    _piv["tipo_dif_rete"] = _piv.apply(_clr, axis=1)
    _piv["tipo_dif_iva"]  = _piv.apply(
        lambda r: "OK" if r["dif_iva"]==0 else ("SIN TC14" if r["neto_14"]==0 else "DIFERENCIA PARCIAL"),
        axis=1)

    df_iva_dif  = _piv[_piv["dif_iva"]!=0].sort_values("inm_num").reset_index()
    df_rete_dif = _piv[_piv["dif_rete"]!=0].sort_values("inm_num").reset_index()
    resumen_rete = _piv.groupby("tipo_dif_rete").agg(
        inmuebles=("dif_rete","count"), dif_total=("dif_rete","sum")
    ).reset_index()

    analisis = dict(
        df_iva_dif=df_iva_dif, df_rete_dif=df_rete_dif, resumen_rete=resumen_rete,
        n_sin_tc23=len(_piv[_piv["tipo_dif_rete"]=="SIN TC23"]),
        n_signo_inv=len(_piv[_piv["tipo_dif_rete"]=="SIGNO INVERTIDO"]),
        n_sin_tc14=len(_piv[_piv["tipo_dif_iva"]=="SIN TC14"]),
        sum_sin_tc23=int(_piv[_piv["tipo_dif_rete"]=="SIN TC23"]["rete_calc"].sum()),
    )

    return df_v1, df_v2, df_v3, crosscheck, analisis, prop_inm, cont_inm_owners


# ── Generar Excel ──────────────────────────────────────────────────────────────
def build_excel(df_v1, df_v2, df_v3, crosscheck, analisis):
    wb  = openpyxl.Workbook()

    # ── Resumen ──────────────────────────────────────────────────────────────
    ws  = wb.active; ws.title = "Resumen"
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:F1")
    t = ws["A1"]
    t.value = "REPORTE DE CONCILIACIÓN RM2 — CORTE ABRIL 2026"
    t.font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor=C_HEADER_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A2:F2")
    ws["A2"].value = "Solo propietarios activos (Ex Prop = False)"
    ws["A2"].font  = Font(name="Arial", size=9, italic=True, color="888888")
    ws["A2"].alignment = Alignment(horizontal="center")

    FMT_NUM   = '#,##0'
    total_inm = df_v1["Inmueble"].nunique()
    correctos = len(df_v1[df_v1["Estado"]=="CORRECTO"])
    al_nombre = len(df_v1[df_v1["Estado"]=="ALERTA — NOMBRE DIFIERE"])
    al_nit    = len(df_v1[df_v1["Estado"]=="ALERTA — NIT NO COINCIDE"])
    solo_p    = len(df_v1[df_v1["Estado"]=="SOLO EN PROPIETARIOS"])
    solo_c    = len(df_v1[df_v1["Estado"]=="SOLO EN CONTABILIDAD"])
    adicional = len(df_v1[df_v1["Estado"]=="PROPIETARIO ADICIONAL"])
    tot_v2    = len(df_v2)
    cor_v2    = len(df_v2[df_v2["Estado"]=="CORRECTO"])
    rev_v2    = len(df_v2[df_v2["Estado"]=="REVISAR"])
    sin_v2    = len(df_v2[df_v2["Estado"]=="CÓDIGO NO EXISTE"])
    tot_v3    = len(df_v3)
    iva_v3    = len(df_v3[df_v3["tipo_causa"]=="3"])
    rete_v3   = len(df_v3[df_v3["tipo_causa"]=="4"])
    deb_v3    = int(df_v3["debito"].sum())
    cred_v3   = int(df_v3["credito"].sum())
    neto_v3   = cred_v3 - deb_v3
    cc        = crosscheck

    def section(start, title, rows):
        ws.cell(start, 1, title).font = Font(name="Arial", size=11, bold=True, color=C_HEADER_BG)
        ws.cell(start+1, 1, "Métrica"); ws.cell(start+1, 2, "Valor")
        _hdr(ws, start+1, 2)
        for off, (lbl, val, bg, fg) in enumerate(rows):
            r = start + 2 + off
            ws.cell(r, 1, lbl).font = Font(name="Arial", size=9)
            cv = ws.cell(r, 2, val); cv.alignment = Alignment(horizontal="center")
            _row(ws, r, 2, C_ALT if r%2==0 else None)
            if bg: _color(cv, bg, fg)

    def crosscheck_section(start, title, rows):
        t2 = ws.cell(start, 1, title)
        t2.font = Font(name="Arial", size=11, bold=True, color=C_HEADER_BG)
        ws.merge_cells(start_row=start, start_column=1, end_row=start, end_column=4)
        for c, h in enumerate(["Concepto","Valor","","Observación"], 1):
            ws.cell(start+1, c, h)
        _hdr(ws, start+1, 4)
        for off, (lbl, val, num_fmt, bg, fg) in enumerate(rows):
            r = start + 2 + off
            ws.cell(r, 1, lbl).font = Font(name="Arial", size=9)
            cv = ws.cell(r, 2, val)
            cv.font = Font(name="Arial", size=9, bold=(bg is not None))
            cv.alignment = Alignment(horizontal="right")
            if num_fmt: cv.number_format = num_fmt
            _row(ws, r, 4, C_ALT if r%2==0 else None)
            if bg:
                for c in [1, 2, 4]:
                    cell = ws.cell(r, c)
                    cell.fill = PatternFill("solid", fgColor=bg)
                    cell.font = Font(name="Arial", size=9, bold=True, color=fg)

    section(4, "VERIFICACIÓN 1 — PROPIETARIOS POR INMUEBLE", [
        ("Total inmuebles verificados",                    total_inm, None,        "000000"),
        ("✓ Correctos (nombre e ID coinciden)",            correctos, C_CORRECT,   C_CORRECT_FG),
        ("⚠ Alerta — nombre difiere",                     al_nombre, C_WARN,      C_WARN_FG),
        ("✗ Alerta — NIT no coincide",                    al_nit,    C_ERROR,     C_ERROR_FG),
        ("→ Solo en Propietarios (sin movimientos)",       solo_p,    C_ONLY_PROP, C_ONLY_PROP_FG),
        ("→ Solo en Contabilidad (sin propietario activo)",solo_c,    C_ONLY_CONT, C_ONLY_CONT_FG),
        ("ℹ Propietarios adicionales en Contabilidad",    adicional, None,        "888888"),
    ])
    section(15, "VERIFICACIÓN 2 — TIPO CAUSA", [
        ("Total filas verificadas",                        tot_v2, None,      "000000"),
        ("✓ Correctos",                                    cor_v2, C_CORRECT, C_CORRECT_FG),
        ("⚠ Para revisar",                                rev_v2, C_WARN,    C_WARN_FG),
        ("✗ Código no existe en catálogo",                 sin_v2, C_ERROR,   C_ERROR_FG),
    ])
    section(23, "VERIFICACIÓN 3 — IVA Y RETENCIONES (TC 3 y TC 4)", [
        ("Total registros IVA/Retención (TC 3 y 4)", tot_v3,  None,  "000000"),
        ("  TC 3 — IVA Arrendamiento 19%",           iva_v3,  None,  "000000"),
        ("  TC 4 — Retención en la Fuente Canon",    rete_v3, None,  "000000"),
        ("Total Débito",                             deb_v3,  None,  "000000"),
        ("Total Crédito",                            cred_v3, None,  "000000"),
        ("Neto (Crédito - Débito)",                  neto_v3,
         C_ERROR if neto_v3 < 0 else C_CORRECT,
         C_ERROR_FG if neto_v3 < 0 else C_CORRECT_FG),
    ])
    for off in [3, 4, 5]:
        ws.cell(23 + 2 + off, 2).number_format = FMT_NUM

    crosscheck_section(33, "VERIFICACIÓN CRUZADA IVA — TC20 × 19% debe = TC14", [
        ("Neto TC 20 (Comisión Arrendamiento)",       cc["neto_20"],   FMT_NUM, None,     "000000"),
        ("× IVA 19% → valor calculado",               cc["iva_calc"],  FMT_NUM, None,     "000000"),
        ("Neto TC 14 (IVA 19% Comisión) — real",      cc["neto_14"],   FMT_NUM, None,     "000000"),
        ("Diferencia (Calculado − Real)",              cc["dif_iva"],   FMT_NUM,
         None if cc["ok_iva"] else C_ERROR, C_CORRECT_FG if cc["ok_iva"] else C_ERROR_FG),
        ("Estado",
         "✓ CUADRA" if cc["ok_iva"] else f"✗ DIFIERE en {cc['dif_iva']:,}",
         None,
         C_CORRECT if cc["ok_iva"] else C_ERROR, C_CORRECT_FG if cc["ok_iva"] else C_ERROR_FG),
    ])
    crosscheck_section(42, "VERIFICACIÓN CRUZADA RETEFUENTE — TC20 × 11% debe = TC23", [
        ("Neto TC 20 (Comisión Arrendamiento)",       cc["neto_20"],    FMT_NUM, None,     "000000"),
        ("× Retefuente 11% → valor calculado",        cc["rete_calc"],  FMT_NUM, None,     "000000"),
        ("Neto TC 23 (Retención Fuente Comisión) — real", cc["neto_23"], FMT_NUM, None,   "000000"),
        ("Diferencia (Calculado − Real)",              cc["dif_rete"],   FMT_NUM,
         None if cc["ok_rete"] else C_ERROR, C_CORRECT_FG if cc["ok_rete"] else C_ERROR_FG),
        ("Estado",
         "✓ CUADRA" if cc["ok_rete"] else f"✗ DIFIERE en {cc['dif_rete']:,}",
         None,
         C_CORRECT if cc["ok_rete"] else C_ERROR, C_CORRECT_FG if cc["ok_rete"] else C_ERROR_FG),
    ])
    # ── Diagnóstico ──────────────────────────────────────────────────────────
    an = analisis
    ws.cell(51, 1, "DIAGNÓSTICO DE DIFERENCIAS").font = Font(
        name="Arial", size=11, bold=True, color=C_HEADER_BG)
    ws.cell(52, 1, "Diagnóstico"); ws.cell(52, 2, "Cantidad")
    _hdr(ws, 52, 2)
    diag_rows = [
        ("IVA — Inmuebles sin TC14",                         an["n_sin_tc14"],   C_ERROR if an["n_sin_tc14"]  else C_CORRECT, C_ERROR_FG if an["n_sin_tc14"]  else C_CORRECT_FG),
        ("Retefuente — Inmuebles sin TC23",                  an["n_sin_tc23"],   C_ERROR if an["n_sin_tc23"]  else C_CORRECT, C_ERROR_FG if an["n_sin_tc23"]  else C_CORRECT_FG),
        ("Retefuente — TC23 signo invertido",                an["n_signo_inv"],  C_WARN  if an["n_signo_inv"] else C_CORRECT, C_WARN_FG  if an["n_signo_inv"] else C_CORRECT_FG),
        ("Retefuente faltante acumulada (sin TC23)",         an["sum_sin_tc23"], C_ERROR if an["sum_sin_tc23"] else C_CORRECT, C_ERROR_FG if an["sum_sin_tc23"] else C_CORRECT_FG),
    ]
    for off, (lbl, val, bg, fg) in enumerate(diag_rows):
        r = 53 + off
        ws.cell(r, 1, lbl).font = Font(name="Arial", size=9)
        cv = ws.cell(r, 2, val); cv.alignment = Alignment(horizontal="center")
        cv.number_format = FMT_NUM
        _row(ws, r, 2, C_ALT if r%2==0 else None)
        _color(cv, bg, fg)

    _widths(ws, [55, 18, 2, 30])

    # ── Ver4 Análisis de Diferencias ─────────────────────────────────────────
    ws4 = wb.create_sheet("Ver4 Análisis Diferencias")
    ws4.sheet_view.showGridLines = False

    def v4t(row, text, ncols):
        ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
        t2 = ws4.cell(row, 1, text)
        t2.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        t2.fill = PatternFill("solid", fgColor=C_HEADER_BG)
        t2.alignment = Alignment(horizontal="center", vertical="center")
        ws4.row_dimensions[row].height = 22

    FMT_NUM2 = '#,##0'
    dcols = ["No. Inm","Nombre Tercero","Neto TC20","Calculado","Neto Real","Diferencia","Diagnóstico"]

    # Sección IVA
    v4t(1, "DIFERENCIAS IVA — Inmuebles con TC20 sin TC14", 7)
    for c, h in enumerate(dcols, 1): ws4.cell(2, c, h)
    _hdr(ws4, 2, 7)
    df_id = an["df_iva_dif"]
    if len(df_id) == 0:
        ws4.merge_cells(start_row=3, start_column=1, end_row=3, end_column=7)
        c2 = ws4.cell(3, 1, "✓ Sin diferencias de IVA")
        c2.font = Font(name="Arial", size=9, color=C_CORRECT_FG)
        c2.fill = PatternFill("solid", fgColor=C_CORRECT)
    else:
        for ri, rw in df_id.iterrows():
            r = ri + 3
            for ci, v in enumerate([rw["inm_num"],rw["nombre"],rw["neto_20"],rw["iva_calc"],rw["neto_14"],rw["dif_iva"],rw["tipo_dif_iva"]], 1):
                ws4.cell(r, ci, v)
            _row(ws4, r, 7, C_ALT if ri%2==0 else None)
            for ci in [3,4,5,6]: ws4.cell(r,ci).number_format=FMT_NUM2; ws4.cell(r,ci).alignment=Alignment(horizontal="right")
            _color(ws4.cell(r,7), C_ERROR, C_ERROR_FG)

    sep = len(df_id) + 5
    # Resumen retefuente
    v4t(sep, "RETEFUENTE — Resumen por categoría", 7)
    for c, h in enumerate(["Categoría","Inmuebles","Diferencia","Interpretación"], 1): ws4.cell(sep+1, c, h)
    _hdr(ws4, sep+1, 4)
    cats = {"SIN TC23":(C_ERROR,C_ERROR_FG,"TC23 no registrado"),"SIGNO INVERTIDO":(C_WARN,C_WARN_FG,"TC23 como crédito — posible convención contable"),"DIFERENCIA PARCIAL":(C_ERROR,C_ERROR_FG,"Requiere revisión")}
    _rr = sep + 2
    for _, rw in an["resumen_rete"].iterrows():
        cat = rw["tipo_dif_rete"]
        if cat == "OK": continue
        bg2, fg2, interp = cats.get(cat, (None,"000000",""))
        ws4.cell(_rr,1,cat); ws4.cell(_rr,2,int(rw["inmuebles"])).alignment=Alignment(horizontal="center")
        ws4.cell(_rr,3,int(rw["dif_total"])).number_format=FMT_NUM2; ws4.cell(_rr,3).alignment=Alignment(horizontal="right")
        ws4.cell(_rr,4,interp); _row(ws4,_rr,4,C_ALT if _rr%2==0 else None)
        if bg2:
            for ci in [1,2,3]: _color(ws4.cell(_rr,ci),bg2,fg2)
        _rr += 1

    sep2 = _rr + 2
    # Sin TC23
    v4t(sep2, "DETALLE — Inmuebles SIN TC23 (retefuente no registrada)", 7)
    for c, h in enumerate(dcols, 1): ws4.cell(sep2+1, c, h)
    _hdr(ws4, sep2+1, 7)
    df_sc = an["df_rete_dif"][an["df_rete_dif"]["tipo_dif_rete"]=="SIN TC23"].reset_index(drop=True)
    for ri, rw in df_sc.iterrows():
        r = sep2 + 2 + ri
        for ci, v in enumerate([rw["inm_num"],rw["nombre"],rw["neto_20"],rw["rete_calc"],rw["neto_23"],rw["dif_rete"],"TC23 NO REGISTRADO"], 1):
            ws4.cell(r, ci, v)
        _row(ws4, r, 7, C_ALT if ri%2==0 else None)
        for ci in [3,4,5,6]: ws4.cell(r,ci).number_format=FMT_NUM2; ws4.cell(r,ci).alignment=Alignment(horizontal="right")
        _color(ws4.cell(r,7), C_ERROR, C_ERROR_FG)

    sep3 = sep2 + 2 + len(df_sc) + 2
    # Signo invertido
    v4t(sep3, "DETALLE — TC23 SIGNO INVERTIDO (revisar convención contable)", 7)
    for c, h in enumerate(dcols, 1): ws4.cell(sep3+1, c, h)
    _hdr(ws4, sep3+1, 7)
    df_si = an["df_rete_dif"][an["df_rete_dif"]["tipo_dif_rete"]=="SIGNO INVERTIDO"].reset_index(drop=True)
    for ri, rw in df_si.iterrows():
        r = sep3 + 2 + ri
        for ci, v in enumerate([rw["inm_num"],rw["nombre"],rw["neto_20"],rw["rete_calc"],rw["neto_23"],rw["dif_rete"],"SIGNO INVERTIDO"], 1):
            ws4.cell(r, ci, v)
        _row(ws4, r, 7, C_ALT if ri%2==0 else None)
        for ci in [3,4,5,6]: ws4.cell(r,ci).number_format=FMT_NUM2; ws4.cell(r,ci).alignment=Alignment(horizontal="right")
        _color(ws4.cell(r,7), C_WARN, C_WARN_FG)

    _widths(ws4, [12, 35, 16, 20, 16, 16, 30])
    ws4.freeze_panes = "A3"

    # ── Ver1 Propietarios ─────────────────────────────────────────────────────
    ws1 = wb.create_sheet("Ver1 Propietarios")
    ws1.sheet_view.showGridLines = False
    ws1.merge_cells("A1:G1")
    t = ws1["A1"]
    t.value = "VERIFICACIÓN 1 — Propietarios por Inmueble (solo activos)"
    t.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor=C_HEADER_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 25

    cols1 = ["No. Inmueble","Propietario (Propietarios)","Cédula/NIT (Prop.)",
             "Propietario (Contabilidad)","NIT (Contabilidad)","Estado","⚠ Alerta"]
    src1  = ["Inmueble","Propietario_Prop","NIT_Prop",
             "Propietario_Cont","NIT_Cont","Estado","Alerta"]
    for c, h in enumerate(cols1, 1): ws1.cell(2, c, h)
    _hdr(ws1, 2, len(cols1))

    ec_col = 6
    al_col = 7
    for ri, row in df_v1.iterrows():
        r = ri + 3
        for ci, s in enumerate(src1, 1): ws1.cell(r, ci, row[s])
        _row(ws1, r, len(cols1), C_ALT if ri%2==0 else None)
        ec = ws1.cell(r, ec_col)
        st = row["Estado"]
        if st == "CORRECTO":
            _color(ec, C_CORRECT, C_CORRECT_FG)
        elif "ALERTA" in st:
            _color(ec, C_ERROR, C_ERROR_FG)
            _color(ws1.cell(r, al_col), C_ERROR, C_ERROR_FG, bold=False)
        elif st == "SOLO EN PROPIETARIOS":
            _color(ec, C_ONLY_PROP, C_ONLY_PROP_FG)
        elif st == "SOLO EN CONTABILIDAD":
            _color(ec, C_ONLY_CONT, C_ONLY_CONT_FG)
        elif st == "PROPIETARIO ADICIONAL":
            _color(ec, C_WARN, C_WARN_FG)

    _widths(ws1, [13,33,18,33,18,26,55])
    ws1.freeze_panes = "A3"

    # ── Ver2 TipoCausa ────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Ver2 TipoCausa")
    ws2.sheet_view.showGridLines = False
    ws2.merge_cells("A1:I1")
    t = ws2["A1"]
    t.value = "VERIFICACIÓN 2 — Tipo Causa vs Detalle (ordenado por No. Inmueble)"
    t.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor=C_HEADER_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 25

    cols2 = ["Fila","No. Inmueble","NIT","Nombre","TipoCausa",
             "Concepto (Catálogo)","Detalle Contabilidad","Estado","Observación"]
    src2  = ["Fila","Inmueble","NIT","Nombre","TipoCausa",
             "Concepto","Detalle","Estado","Observacion"]
    for c, h in enumerate(cols2, 1): ws2.cell(2, c, h)
    _hdr(ws2, 2, len(cols2))

    ec2 = cols2.index("Estado") + 1
    for ri, row in df_v2.iterrows():
        r = ri + 3
        for ci, s in enumerate(src2, 1): ws2.cell(r, ci, row[s])
        _row(ws2, r, len(cols2), C_ALT if ri%2==0 else None)
        ec = ws2.cell(r, ec2)
        if row["Estado"] == "CORRECTO":
            _color(ec, C_CORRECT, C_CORRECT_FG)
        elif row["Estado"] == "REVISAR":
            _color(ec, C_WARN, C_WARN_FG)
        elif row["Estado"] in ("CÓDIGO NO EXISTE","SIN KEYWORDS"):
            _color(ec, C_ERROR, C_ERROR_FG)

    _widths(ws2, [6,13,14,30,11,30,60,15,45])
    ws2.freeze_panes = "A3"

    # ── Ver3 IVA y Retenciones ────────────────────────────────────────────────
    ws3 = wb.create_sheet("Ver3 IVA y Retenciones")
    ws3.sheet_view.showGridLines = False
    ws3.merge_cells("A1:H1")
    t = ws3["A1"]
    t.value = "VERIFICACIÓN 3 — IVA (TC 3) y Retención en la Fuente (TC 4)"
    t.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor=C_HEADER_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 25

    cols3 = ["No. Inmueble","Nombre Tercero","TC","Concepto","Detalle","Débito","Crédito","Neto (Cred - Deb)"]
    for c, h in enumerate(cols3, 1): ws3.cell(2, c, h)
    _hdr(ws3, 2, len(cols3))

    C_IVA  = "DEEAF1"
    C_RETE = "EBF0DE"

    for ri, row in df_v3.iterrows():
        r = ri + 3
        neto = int(row["credito"]) - int(row["debito"])
        vals = [row["inm_num"], row["nombre_cont"], row["tipo_causa"],
                row["concepto"], row["detalle"], int(row["debito"]), int(row["credito"]), neto]
        for ci, val in enumerate(vals, 1): ws3.cell(r, ci, val)
        bg = C_IVA if row["tipo_causa"] == "3" else C_RETE
        _row(ws3, r, len(cols3), bg)
        for ci in [1, 6, 7, 8]:
            ws3.cell(r, ci).alignment = Alignment(horizontal="right")
        if neto < 0:
            ws3.cell(r, 8).font = Font(name="Arial", size=9, bold=True, color=C_ERROR_FG)

    r_tot = len(df_v3) + 3
    ws3.cell(r_tot, 5, "TOTAL").font = Font(name="Arial", size=9, bold=True)
    tot_deb  = int(df_v3["debito"].sum())
    tot_cred = int(df_v3["credito"].sum())
    tot_neto = tot_cred - tot_deb
    for ci, val in [(6, tot_deb), (7, tot_cred), (8, tot_neto)]:
        cell = ws3.cell(r_tot, ci, val)
        cell.font = Font(name="Arial", size=9, bold=True)
        cell.fill = PatternFill("solid", fgColor="D9D9D9")
        cell.alignment = Alignment(horizontal="right")
        cell.border = Border(top=Side(style="thin", color="000000"))

    _widths(ws3, [14,32,5,32,65,14,14,18])
    ws3.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════════════════
# INTERFAZ STREAMLIT
# ══════════════════════════════════════════════════════════════════════════════


# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO: PLANTILLA CONTAI
# ══════════════════════════════════════════════════════════════════════════════
import io
import re
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# ── Constantes ─────────────────────────────────────────────────────────────────
NIT_FONDO     = 901155077
CENTRO_COSTO  = "0501"
CUENTA_CONTRA = 22050501

C_HEADER_BG = "1F4E79"
C_HEADER_FG = "FFFFFF"
C_ALT       = "F2F2F2"

CUENTAS = {
    1:  {"gravado": 41550502, "no_gravado": 41550501},
    2:  61053502,
    3:  24080502,
    4:  13551513,
    7:  61053501,
    11: 61053005,
    14: 24081002,
    20: 61950502,
    23: 23652007,
    30: 61055001,
    32: 61051501,
    33: 51350505,
    37: 42505002,
}

CUENTAS_REVERSION = {
    1:  {"gravado": 41750501, "no_gravado": 41750502},
    3:  24080597,
    4:  13551513,
    14: 24081075,
    20: 61970505,
    23: 23652091,
}

TC_ACTA       = {1, 2, 3, 4, 7, 11, 30, 32, 33, 37}
TC_COMISIONES = {20, 14, 23}
TC_EXCLUIR    = {0, 16, 40}

MESES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO",
    6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE",
    10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}

NOMBRE_CUENTAS = {
    41550502: "Canon Gravado",       41550501: "Canon No Gravado",
    61053502: "Cuota Admon",         24080502: "IVA Comercial",
    13551513: "Retefte",             61053501: "EPM",
    61053005: "Seguro",              24081002: "IVA Comisión",
    61950502: "Comisión",            23652007: "Retefte Comisión",
    61055001: "Rep. Locativas",      61051501: "Prediales",
    51350505: "Aseo",                42505002: "Apropi. Depósito",
    41750501: "Rev. Canon Gravado",  41750502: "Rev. Canon Excluido",
    24080597: "Rev. IVA Comercial",  24081075: "Rev. IVA Comisión",
    61970505: "Rev. Comisión",       23652091: "Rev. Retefte Comisión",
    22050501: "Contrapartida Comisión",
}

NOMBRE_TC = {
    1: "Canon", 2: "Cuota Admon", 3: "IVA Comercial", 4: "Retefte",
    7: "EPM", 11: "Seguro", 14: "IVA Comisión", 20: "Comisión",
    23: "Retefte Comisión", 30: "Rep. Locativas", 32: "Prediales",
    33: "Aseo", 37: "Apropi. Depósito",
}


# ── Helpers ────────────────────────────────────────────────────────────────────
def parse_valor(val):
    if pd.isna(val):
        return 0.0
    clean = re.sub(r"[$ ,]", "", str(val)).strip()
    try:
        return float(clean)
    except ValueError:
        return 0.0


def es_reversion(detalle):
    return "revers" in str(detalle).lower() if pd.notna(detalle) else False


def format_fecha(fecha):
    if pd.isna(fecha):
        return ""
    if isinstance(fecha, str):
        parts = fecha.strip().split("/")
        if len(parts) == 3:
            return f"{parts[1]}/{parts[0]}/{parts[2]}"
        return fecha
    return fecha.strftime("%m/%d/%Y")


def get_mes(fecha):
    if pd.isna(fecha):
        return None
    if isinstance(fecha, str):
        parts = fecha.strip().split("/")
        if len(parts) == 3:
            try:
                return int(parts[1])
            except ValueError:
                return None
    try:
        return fecha.month
    except Exception:
        return None


def format_ad_ref(trans_int):
    """2026030005 → 'AD2630005'  (AD + año_corto + mes_sin_cero + secuencia)"""
    s = str(trans_int)
    if len(s) == 10:
        year_short = s[2:4]
        month = str(int(s[4:6]))
        seq = s[6:]
        return f"AD{year_short}{month}{seq}"
    return f"AD{trans_int}"


def get_cuenta(tc, reversion, tiene_iva):
    tabla = CUENTAS_REVERSION if reversion and tc in CUENTAS_REVERSION else CUENTAS
    entrada = tabla.get(tc)
    if entrada is None:
        return None
    if isinstance(entrada, dict):
        # Forward: CON IVA → No Gravado (41550501). SIN IVA → Gravado (41550502)
        # Reversal: CON IVA → Gravado reversal (41750501). SIN IVA → No Gravado reversal (41750502)
        if reversion:
            return entrada["gravado"] if tiene_iva else entrada["no_gravado"]
        else:
            return entrada["no_gravado"] if tiene_iva else entrada["gravado"]
    return entrada


def tipo_mov(debito, credito):
    """Crédito en fuente → 2 en planilla. Débito en fuente → 1."""
    return 2 if credito > 0 else 1


def make_row(cuenta, comp, fecha, doc, doc_ref, detalle, tipo, valor, base):
    return {
        "Cuenta":            cuenta,
        "Comprobante":       comp,
        "Fecha(mm/dd/yyyy)": fecha,
        "Documento":         doc,
        "Documento Ref.":    doc_ref,
        "Nit":               NIT_FONDO,
        "Detalle":           str(detalle) if pd.notna(detalle) else "",
        "Tipo":              tipo,
        "Valor":             valor,
        "Base":              base if pd.notna(base) else "",
        "Centro de Costo":   CENTRO_COSTO,
    }


# ── Procesador ACTA ────────────────────────────────────────────────────────────
def procesar_acta(df_src, documento, fecha_acta):
    """
    Un único Documento y una única fecha para todas las filas del ACTA.
    Siempre CP. Orden por No. Transaccion ASC, luego TipoCausa ASC.
    Gravado (41550502) = TC=1 SIN TC=3 paired.
    No Gravado (41550501) = TC=1 CON TC=3 paired.
    """
    trans_con_iva = set(
        df_src.loc[df_src["TipoCausa"] == 3, "No. Transaccion"]
        .dropna().astype(int).tolist()
    )

    filas = df_src[
        df_src["TipoCausa"].notna() &
        df_src["TipoCausa"].isin(TC_ACTA) &
        df_src["No. Transaccion"].notna()
    ].copy()
    filas["TipoCausa"] = filas["TipoCausa"].astype(int)
    filas["_trans"]    = filas["No. Transaccion"].astype(int)

    # Orden: primero TC=1/3/4 (Canon+IVA+Retefte) juntos por No.Trans ASC,
    # luego el resto de TCs por No.Trans ASC — dentro de cada grupo por TipoCausa ASC
    TC_GRUPO1 = {1, 3, 4}
    filas["_grupo"] = filas["TipoCausa"].apply(lambda t: 0 if t in TC_GRUPO1 else 1)
    filas = filas.sort_values(["_grupo", "_trans", "TipoCausa"], ascending=[True, True, True])

    resultado = []
    for _, row in filas.iterrows():
        tc       = int(row["TipoCausa"])
        detalle  = row["Detalle"]
        rev      = es_reversion(detalle)
        trans_id = int(row["_trans"])
        tiene_iva = trans_id in trans_con_iva

        cuenta = get_cuenta(tc, rev, tiene_iva)
        if cuenta is None:
            continue

        deb   = parse_valor(row["Debito"])
        cred  = parse_valor(row["Credito"])
        base  = parse_valor(row["Valor Base"])
        tipo  = tipo_mov(deb, cred)
        valor = cred if cred > 0 else deb

        resultado.append(make_row(
            cuenta, "CP", fecha_acta,      # siempre CP, siempre fecha_acta
            documento, documento,
            detalle, tipo, valor, base,
        ))

    return resultado


# ── Procesador COMISIONES ──────────────────────────────────────────────────────
def procesar_comisiones(df_src, cons_cp, cons_dv):
    """
    Agrupa por No. Transaccion. Cada grupo = una factura.
    Agrega fila de contrapartida 22050501 por grupo.
    CP (regulares) y DV (reversiones) tienen consecutivos separados.
    """
    filas = df_src[
        df_src["TipoCausa"].notna() &
        df_src["TipoCausa"].isin(TC_COMISIONES) &
        df_src["No. Transaccion"].notna()
    ].copy()
    filas["TipoCausa"] = filas["TipoCausa"].astype(int)
    filas["_rev"]      = filas["Detalle"].apply(es_reversion)

    resultado = []
    seq_cp = cons_cp
    seq_dv = cons_dv

    for trans_id, grupo in filas.groupby("No. Transaccion", sort=True):
        rev   = bool(grupo["_rev"].iloc[0])
        comp  = "DV" if rev else "CP"
        doc   = seq_dv if rev else seq_cp
        trans_int = int(trans_id)

        # Doc Ref: "FE" + No. Transaccion para CP; formato AD especial para DV
        doc_ref = format_ad_ref(trans_int) if rev else f"FE{trans_int}"

        fecha_str  = format_fecha(grupo.iloc[0]["Fecha"])
        mes        = get_mes(grupo.iloc[0]["Fecha"])
        mes_nombre = MESES.get(mes, "")

        sum_com = sum_iva = sum_ret = 0.0

        for _, row in grupo.iterrows():
            tc     = int(row["TipoCausa"])
            deb    = parse_valor(row["Debito"])
            cred   = parse_valor(row["Credito"])
            base   = parse_valor(row["Valor Base"]) if tc != 20 else ""  # TC=20 base vacío
            tipo   = tipo_mov(deb, cred)
            valor  = cred if cred > 0 else deb
            cuenta = get_cuenta(tc, rev, False)
            if cuenta is None:
                continue

            if tc == 20:
                sum_com += valor
            elif tc == 14:
                sum_iva += valor
            elif tc == 23:
                sum_ret += valor

            resultado.append(make_row(
                cuenta, comp, fecha_str, doc, doc_ref,
                row["Detalle"], tipo, valor, base,
            ))

        # Contrapartida 22050501
        valor_contra = sum_com + sum_iva - sum_ret
        # Tipo opuesto al de la comisión:
        # CP → TC=20 es DEBITO → tipo=1 → contra es 2
        # DV → TC=20 es CREDITO → tipo=2 → contra es 1
        tipo_contra = 1 if rev else 2
        resultado.append(make_row(
            CUENTA_CONTRA, comp, fecha_str, doc, doc_ref,
            f"COMISION MES {mes_nombre}", tipo_contra, valor_contra, "",
        ))

        if rev:
            seq_dv += 1
        else:
            seq_cp += 1

    return resultado


# ── Generar Excel ──────────────────────────────────────────────────────────────
COLS   = ["Cuenta", "Comprobante", "Fecha(mm/dd/yyyy)", "Documento",
          "Documento Ref.", "Nit", "Detalle", "Tipo", "Valor", "Base",
          "Centro de Costo"]
WIDTHS = [12, 13, 17, 13, 14, 13, 60, 6, 14, 10, 14]


def _hdr(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(1, c)
        cell.font      = Font(bold=True, color=C_HEADER_FG, name="Arial", size=10)
        cell.fill      = PatternFill("solid", fgColor=C_HEADER_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = Border(bottom=Side(style="thin", color="000000"),
                                right=Side(style="thin", color="CCCCCC"))


def _fila(ws, row_num, ncols, alt):
    for c in range(1, ncols + 1):
        cell = ws.cell(row_num, c)
        cell.font      = Font(name="Arial", size=9)
        cell.alignment = Alignment(vertical="center")
        cell.border    = Border(bottom=Side(style="thin", color="E0E0E0"),
                                right=Side(style="thin", color="E0E0E0"))
        if alt:
            cell.fill = PatternFill("solid", fgColor=C_ALT)


def generar_excel(rows, sheet_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.freeze_panes = "A2"

    for c, name in enumerate(COLS, 1):
        ws.cell(1, c, name)
    _hdr(ws, len(COLS))
    ws.row_dimensions[1].height = 30

    for i, row in enumerate(rows, 2):
        _fila(ws, i, len(COLS), i % 2 == 0)
        for c, key in enumerate(COLS, 1):
            v    = row.get(key, "")
            cell = ws.cell(i, c, v if v is not None else "")
            if c in (1, 4, 6):                                     # Cuenta, Documento, Nit
                cell.number_format = "0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c == 5:                                            # Documento Ref.
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in (9, 10) and isinstance(v, (int, float)):     # Valor, Base
                cell.number_format = '#,##0;(#,##0);"-"'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c in (2, 3, 8):                                    # Comprobante, Fecha, Tipo
                cell.alignment = Alignment(horizontal="center", vertical="center")

    for c, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[get_column_letter(c)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def resumen_cuentas(rows):
    df = pd.DataFrame(rows)
    res = (df.groupby("Cuenta")
             .agg(Filas=("Valor", "count"), Total=("Valor", "sum"))
             .reset_index())
    res["Nombre"] = res["Cuenta"].map(NOMBRE_CUENTAS).fillna("—")
    res["Total"]  = res["Total"].apply(lambda x: f"${x:,.0f}")
    return res[["Cuenta", "Nombre", "Filas", "Total"]]




import datetime

# ══════════════════════════════════════════════════════════════════════════════
# APP COMBINADA — CONTABILIDAD RM2
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Contabilidad RM2",
    page_icon="🏢",
    layout="wide",
)

# ── CSS global ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background: #f0f2f6; }
  [data-testid="stSidebar"]          { display: none; }

  /* Banner principal */
  .main-banner {
    background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
    border-radius: 14px;
    padding: 30px 40px;
    margin-bottom: 28px;
    display: flex;
    align-items: center;
    gap: 20px;
  }
  .main-banner-icon  { font-size: 52px; line-height: 1; }
  .main-banner-title { font-size: 32px; font-weight: 800; color: #fff;
                       margin: 0; letter-spacing: -0.5px; }
  .main-banner-sub   { font-size: 14px; color: rgba(255,255,255,0.75);
                       margin: 6px 0 0 0; }

  /* Tarjetas de navegación */
  .nav-row   { display: flex; gap: 20px; margin-bottom: 28px; }
  .nav-card  {
    flex: 1; background: #fff; border-radius: 12px;
    padding: 28px 32px; box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    border: 2px solid transparent; cursor: pointer;
    transition: border-color .2s, box-shadow .2s;
  }
  .nav-card:hover { border-color: #2E75B6; box-shadow: 0 4px 16px rgba(46,117,182,.15); }
  .nav-card.active { border-color: #1F4E79; box-shadow: 0 4px 16px rgba(31,78,121,.2); }
  .nav-icon  { font-size: 36px; margin-bottom: 10px; }
  .nav-label { font-size: 20px; font-weight: 700; color: #1F4E79; margin: 0; }
  .nav-desc  { font-size: 13px; color: #6b7280; margin: 6px 0 0 0; }

  /* Tabs */
  [data-testid="stTabs"] > div > div { gap: 8px; }
  [data-testid="stTab"] {
    border-radius: 8px 8px 0 0 !important;
    font-weight: 600 !important; font-size: 14px !important;
    color: #374151 !important;
  }
  [data-testid="stTab"][aria-selected="true"] {
    color: #1F4E79 !important;
    border-bottom: 3px solid #1F4E79 !important;
  }
  [data-testid="stTab"] p,
  [data-testid="stTab"] span { color: inherit !important; }

  /* Métricas */
  .metric-row { display: flex; gap: 16px; margin-bottom: 24px; }
  .metric-card {
    flex: 1; background: #fff; border-radius: 10px;
    padding: 20px 24px; box-shadow: 0 1px 4px rgba(0,0,0,.08);
    border-left: 4px solid #2E75B6;
  }
  .metric-label { font-size: 12px; color: #6b7280; font-weight: 600;
                  text-transform: uppercase; letter-spacing: .5px; }
  .metric-value { font-size: 32px; font-weight: 700; color: #1F4E79;
                  margin: 4px 0 0 0; line-height: 1; }

  /* Cards de sección */
  .upload-card, .config-panel {
    background: #fff; border-radius: 10px; padding: 24px;
    box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-bottom: 20px;
  }
  .section-title {
    font-size: 13px; font-weight: 700; color: #374151;
    text-transform: uppercase; letter-spacing: .6px; margin-bottom: 12px;
  }

  /* Pills */
  .pill { display:inline-block; padding:2px 10px; border-radius:20px;
          font-size:12px; font-weight:600; }
  .pill-blue  { background:#dbeafe; color:#1d4ed8; }
  .pill-green { background:#dcfce7; color:#166534; }

  /* Botones primarios */
  [data-testid="stButton"] > button[kind="primary"] {
    background: linear-gradient(135deg,#1F4E79,#2E75B6) !important;
    color: #fff !important; border: none !important;
    border-radius: 8px !important; font-weight: 600 !important; font-size:15px !important;
  }
  [data-testid="stButton"] > button[kind="primary"]:hover {
    background: linear-gradient(135deg,#163d5f,#1F4E79) !important; color:#fff !important;
  }
  /* Botón descarga */
  [data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg,#1F4E79,#2E75B6) !important;
    color: #fff !important; border: none !important;
    border-radius: 8px !important; font-weight: 600 !important;
    padding: 12px !important; font-size:15px !important;
  }
  [data-testid="stDownloadButton"] > button:hover {
    background: linear-gradient(135deg,#163d5f,#1F4E79) !important; color:#fff !important;
  }
  [data-testid="stNumberInput"] label,
  [data-testid="stDateInput"]   label { color:#374151 !important; font-weight:500 !important; }
</style>
""", unsafe_allow_html=True)

# ── Banner principal ───────────────────────────────────────────────────────────
st.markdown("""
<div class="main-banner">
  <div class="main-banner-icon">🏢</div>
  <div>
    <p class="main-banner-title">Contabilidad RM2</p>
    <p class="main-banner-sub">Herramientas de conciliación y generación de plantillas contables</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Navegación ─────────────────────────────────────────────────────────────────
if "modulo" not in st.session_state:
    st.session_state["modulo"] = None

col_n1, col_n2 = st.columns(2)
with col_n1:
    if st.button("🔍  Conciliación", use_container_width=True,
                 type="primary" if st.session_state["modulo"] == "conciliacion" else "secondary",
                 key="nav_concil"):
        st.session_state["modulo"] = "conciliacion"
        st.rerun()
with col_n2:
    if st.button("📋  Actas Fondo — Contai", use_container_width=True,
                 type="primary" if st.session_state["modulo"] == "contai" else "secondary",
                 key="nav_contai"):
        st.session_state["modulo"] = "contai"
        st.rerun()

modulo = st.session_state["modulo"]

# ── Landing (sin módulo seleccionado) ─────────────────────────────────────────
if modulo is None:
    st.markdown("""
    <div style="text-align:center; padding:60px 0; color:#9ca3af;">
      <div style="font-size:56px">👆</div>
      <p style="font-size:18px; font-weight:600; color:#374151; margin-top:16px;">
        Selecciona un módulo para comenzar
      </p>
      <div style="display:flex; gap:32px; justify-content:center; margin-top:32px;">
        <div style="background:#fff;border-radius:12px;padding:28px 36px;
                    box-shadow:0 2px 8px rgba(0,0,0,.08);max-width:280px;text-align:left;">
          <div style="font-size:32px">🔍</div>
          <p style="font-size:17px;font-weight:700;color:#1F4E79;margin:10px 0 6px;">Conciliación</p>
          <p style="font-size:13px;color:#6b7280;margin:0;">Verifica propietarios, tipo causa, IVA
          y retenciones comparando el archivo de propietarios con contabilidad.</p>
        </div>
        <div style="background:#fff;border-radius:12px;padding:28px 36px;
                    box-shadow:0 2px 8px rgba(0,0,0,.08);max-width:280px;text-align:left;">
          <div style="font-size:32px">📋</div>
          <p style="font-size:17px;font-weight:700;color:#1F4E79;margin:10px 0 6px;">Actas Fondo — Contai</p>
          <p style="font-size:13px;color:#6b7280;margin:0;">Genera las planillas contables de ACTA
          y COMISIONES desde el Movimiento Cta 28150501.</p>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO: CONCILIACIÓN
# ══════════════════════════════════════════════════════════════════════════════
if modulo == "conciliacion":
    st.markdown("### 🔍 Conciliación RM2")
    st.markdown("Sube los tres archivos y haz clic en **Conciliar** para generar el reporte.")

    col1, col2, col3 = st.columns(3)
    with col1:
        f_prop = st.file_uploader("📋 Propietarios RM2", type=["xls","xlsx"],
                                  help="Archivo PROPIETARIOS RM2 CORTE ABRIL 30")
    with col2:
        f_cont = st.file_uploader("📒 Contabilidad RM2", type=["xls","xlsx"],
                                  help="Archivo 28150501CONTABILIDAD RM2")
    with col3:
        f_tc = st.file_uploader("📄 Tipo Causa", type=["xls","xlsx"],
                                help="Archivo TIPO CAUSA")

    st.divider()

    if st.button("▶ Conciliar", type="primary", disabled=not (f_prop and f_cont and f_tc)):
        with st.spinner("Procesando..."):
            df_v1, df_v2, df_v3, crosscheck, analisis, prop_inm, cont_inm = run_conciliacion(f_prop, f_cont, f_tc)

        total_inm = df_v1["Inmueble"].nunique()
        correctos = len(df_v1[df_v1["Estado"]=="CORRECTO"])
        alertas   = len(df_v1[df_v1["Estado"].str.startswith("ALERTA")])
        solo_p    = len(df_v1[df_v1["Estado"]=="SOLO EN PROPIETARIOS"])
        solo_c    = len(df_v1[df_v1["Estado"]=="SOLO EN CONTABILIDAD"])
        total_v3  = len(df_v3)
        iva_v3    = len(df_v3[df_v3["tipo_causa"]=="3"])
        rete_v3   = len(df_v3[df_v3["tipo_causa"]=="4"])
        cc        = crosscheck

        st.subheader("Resumen")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Inmuebles verificados", total_inm)
        m2.metric("✅ Correctos", correctos)
        m3.metric("🚨 Alertas", alertas, delta=alertas if alertas else None, delta_color="inverse")
        m4.metric("Solo Propietarios", solo_p)
        m5.metric("Solo Contabilidad", solo_c)

        c1, c2 = st.columns(2)
        with c1:
            if cc["ok_iva"]:
                st.success(f"✅ IVA cuadra — TC20 × 19% = TC14 ({cc['iva_calc']:,})")
            else:
                st.error(f"❌ IVA NO cuadra — Calculado: {cc['iva_calc']:,} · Real (TC14): {cc['neto_14']:,} · **Diferencia: {cc['dif_iva']:,}**")
        with c2:
            if cc["ok_rete"]:
                st.success(f"✅ Retefuente cuadra — TC20 × 11% = TC23 ({cc['rete_calc']:,})")
            else:
                st.error(f"❌ Retefuente NO cuadra — Calculado: {cc['rete_calc']:,} · Real (TC23): {cc['neto_23']:,} · **Diferencia: {cc['dif_rete']:,}**")

        st.divider()

        tab1, tab2, tab3 = st.tabs([
            "📋 Verificación 1 — Propietarios",
            "📄 Verificación 2 — Tipo Causa",
            "💰 Verificación 3 — IVA y Retenciones",
        ])

        with tab1:
            st.markdown("**Resultado por inmueble** (solo propietarios activos `Ex Prop = False`)")
            def color_estado_v1(val):
                if val == "CORRECTO":               return "background-color:#C6EFCE;color:#276221"
                if "ALERTA" in str(val):            return "background-color:#FFCCCC;color:#9C0006"
                if val == "SOLO EN PROPIETARIOS":   return "background-color:#DAE3F3;color:#1F4E79"
                if val == "SOLO EN CONTABILIDAD":   return "background-color:#FCE4D6;color:#843C0C"
                if val == "PROPIETARIO ADICIONAL":  return "background-color:#FFEB9C;color:#9C6500"
                return ""
            display_v1 = df_v1.rename(columns={"Inmueble":"No. Inm","Propietario_Prop":"Propietario (Prop.)","NIT_Prop":"NIT (Prop.)","Propietario_Cont":"Propietario (Cont.)","NIT_Cont":"NIT (Cont.)"})
            filtro = st.selectbox("Filtrar por estado", ["Todos","CORRECTO","ALERTA — NOMBRE DIFIERE","ALERTA — NIT NO COINCIDE","SOLO EN PROPIETARIOS","SOLO EN CONTABILIDAD","PROPIETARIO ADICIONAL"])
            if filtro != "Todos":
                display_v1 = display_v1[display_v1["Estado"] == filtro]
            st.dataframe(display_v1.style.map(color_estado_v1, subset=["Estado"]), use_container_width=True, hide_index=True, height=500)

        with tab2:
            st.markdown("**Resultado por fila** de Contabilidad (ordenado por inmueble)")
            def color_estado_v2(val):
                if val == "CORRECTO":                    return "background-color:#C6EFCE;color:#276221"
                if val == "REVISAR":                     return "background-color:#FFEB9C;color:#9C6500"
                if val in ("CÓDIGO NO EXISTE","SIN KEYWORDS"): return "background-color:#FFCCCC;color:#9C0006"
                return ""
            display_v2 = df_v2.rename(columns={"Inmueble":"No. Inm","Concepto":"Concepto (Catálogo)","Detalle":"Detalle Contabilidad","Observacion":"Observación"})
            filtro2 = st.selectbox("Filtrar por estado ", ["Todos","CORRECTO","REVISAR","CÓDIGO NO EXISTE"])
            if filtro2 != "Todos":
                display_v2 = display_v2[display_v2["Estado"] == filtro2]
            st.dataframe(display_v2.style.map(color_estado_v2, subset=["Estado"]), use_container_width=True, hide_index=True, height=500)

        with tab3:
            st.markdown(f"**IVA Arrendamiento (TC 3): {iva_v3} registros · Retención Fuente Canon (TC 4): {rete_v3} registros**")
            fa, fb = st.columns(2)
            fa.metric("Total IVA (TC 3)", iva_v3)
            fb.metric("Total Retención (TC 4)", rete_v3)
            filtro3 = st.selectbox("Filtrar por tipo causa", ["Todos","3 — IVA Arrendamiento","4 — Retención en la Fuente"])
            display_v3 = df_v3.rename(columns={"inm_num":"No. Inm","nombre_cont":"Nombre Tercero","tipo_causa":"TC","concepto":"Concepto","detalle":"Detalle","debito":"Débito","credito":"Crédito","neto":"Neto"})[["No. Inm","Nombre Tercero","TC","Concepto","Detalle","Débito","Crédito","Neto"]]
            if filtro3 == "3 — IVA Arrendamiento":      display_v3 = display_v3[display_v3["TC"]=="3"]
            elif filtro3 == "4 — Retención en la Fuente": display_v3 = display_v3[display_v3["TC"]=="4"]
            def color_tc(val):
                if val == "3": return "background-color:#DEEAF1;color:#1F4E79"
                if val == "4": return "background-color:#EBF0DE;color:#375623"
                return ""
            def color_neto(val):
                if isinstance(val,(int,float)) and val < 0: return "color:#9C0006;font-weight:bold"
                return ""
            st.dataframe(display_v3.style.map(color_tc,subset=["TC"]).map(color_neto,subset=["Neto"]), use_container_width=True, hide_index=True, height=500)

        st.divider()
        excel_buf = build_excel(df_v1, df_v2, df_v3, crosscheck, analisis)
        st.download_button("⬇️ Descargar REPORTE_CONCILIACION.xlsx", data=excel_buf,
                           file_name="REPORTE_CONCILIACION.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary")
    else:
        if not (f_prop and f_cont and f_tc):
            st.info("👆 Sube los tres archivos para habilitar el botón **Conciliar**.")

# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO: ACTAS FONDO — CONTAI
# ══════════════════════════════════════════════════════════════════════════════
elif modulo == "contai":
    st.markdown("### 📋 Plantilla Acta RM2")

    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.markdown('<p class="section-title">📂 Archivo de movimientos</p>', unsafe_allow_html=True)
    archivo = st.file_uploader(
        "Arrastra o selecciona · Movimiento Cta 28150501 RM2 -Fondo.xls",
        type=["xls","xlsx"],
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if not archivo:
        st.markdown("""
        <div style="text-align:center;padding:40px 0;color:#9ca3af;">
          <div style="font-size:40px">📄</div>
          <p style="font-size:15px;margin-top:10px;">Carga el archivo de movimientos para comenzar</p>
        </div>""", unsafe_allow_html=True)
        st.stop()

    try:
        df_src = pd.read_excel(archivo, sheet_name=0, header=0)
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        st.stop()

    df_src["TipoCausa"] = pd.to_numeric(df_src["TipoCausa"], errors="coerce")

    n_acta = len(df_src[df_src["TipoCausa"].isin(TC_ACTA) & df_src["No. Transaccion"].notna()])
    n_com  = len(df_src[df_src["TipoCausa"].isin(TC_COMISIONES) & df_src["No. Transaccion"].notna()])
    n_fact = df_src[df_src["TipoCausa"].isin(TC_COMISIONES) & df_src["No. Transaccion"].notna()]["No. Transaccion"].nunique()

    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card" style="border-left-color:#1F4E79">
        <div class="metric-label">Archivo cargado</div>
        <div class="metric-value" style="font-size:16px;margin-top:6px;">✅ {archivo.name}</div>
      </div>
      <div class="metric-card" style="border-left-color:#2E75B6">
        <div class="metric-label">Filas ACTA</div>
        <div class="metric-value">{n_acta:,}</div>
      </div>
      <div class="metric-card" style="border-left-color:#4CAF50">
        <div class="metric-label">Filas Comisiones</div>
        <div class="metric-value">{n_com:,}</div>
      </div>
      <div class="metric-card" style="border-left-color:#FF9800">
        <div class="metric-label">Facturas únicas</div>
        <div class="metric-value">{n_fact:,}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    tab_acta, tab_com = st.tabs(["📋  Planilla ACTA", "💼  Planilla COMISIONES"])

    with tab_acta:
        st.markdown('<div class="config-panel">', unsafe_allow_html=True)
        st.markdown('<p class="section-title">⚙️ Parámetros del ACTA</p>', unsafe_allow_html=True)
        st.markdown('<p style="font-size:13px;color:#6b7280;margin-bottom:16px;">El mismo Documento y Fecha se aplican a <strong>todas</strong> las filas. Canon va primero, Prediales/Admon al final.</p>', unsafe_allow_html=True)
        col_a1, col_a2 = st.columns(2)
        with col_a1:
            doc_acta = st.number_input("Consecutivo Documento", min_value=1, value=26040185, step=1, key="doc_acta")
        with col_a2:
            fecha_acta_dt = st.date_input("Fecha del ACTA", value=datetime.date(2026, 4, 30), key="fecha_acta")
        fecha_acta_str = fecha_acta_dt.strftime("%m/%d/%Y")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div style="display:flex;gap:12px;margin-bottom:12px;"><span class="pill pill-blue">Canon · IVA · Retefte · Admon · EPM · Seguro · Prediales · Aseo · Rep. Locativas</span><span class="pill pill-green">Siempre CP</span></div>', unsafe_allow_html=True)

        if st.button("▶  Generar Planilla ACTA", type="primary", key="btn_acta", use_container_width=True):
            with st.spinner("Procesando ACTA…"):
                rows_acta = procesar_acta(df_src, int(doc_acta), fecha_acta_str)
            if not rows_acta:
                st.warning("No se encontraron movimientos válidos para el ACTA.")
            else:
                st.success(f"✅ **{len(rows_acta):,} filas** generadas — Documento: **{doc_acta}** · Fecha: **{fecha_acta_str}**")
                col_r1, col_r2 = st.columns([1, 2])
                with col_r1:
                    st.markdown("**Resumen por cuenta**")
                    st.dataframe(resumen_cuentas(rows_acta), use_container_width=True, hide_index=True, height=320)
                with col_r2:
                    with st.expander("Vista previa — primeras 50 filas", expanded=True):
                        cols_show = [c for c in pd.DataFrame(rows_acta).columns if not c.startswith("_")]
                        st.dataframe(pd.DataFrame(rows_acta)[cols_show].head(50), use_container_width=True)
                st.markdown("---")
                buf = generar_excel(rows_acta, "ACTA_RM2")
                st.download_button("⬇️  Descargar Plantilla Acta RM2.xlsx", data=buf,
                                   file_name="Plantilla_Acta_RM2.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True, type="primary", key="dl_acta")

    with tab_com:
        st.markdown('<div class="config-panel">', unsafe_allow_html=True)
        st.markdown('<p class="section-title">⚙️ Parámetros de COMISIONES</p>', unsafe_allow_html=True)
        st.markdown('<p style="font-size:13px;color:#6b7280;margin-bottom:16px;">Cada factura recibe su propio consecutivo. Contrapartida <strong>22050501</strong> automática. Reversiones usan comprobante <strong>DV</strong>.</p>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            cons_cp = st.number_input("Consecutivo inicial CP (regulares)", min_value=1, value=26040186, step=1, key="cons_cp")
        with col2:
            cons_dv = st.number_input("Consecutivo inicial DV (reversiones)", min_value=1, value=26040001, step=1, key="cons_dv")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div style="display:flex;gap:12px;margin-bottom:12px;"><span class="pill pill-blue">Comisión · IVA Comisión · Retefte Comisión · Contrapartida 22050501</span><span class="pill pill-green">CP regulares · DV reversiones</span></div>', unsafe_allow_html=True)

        if st.button("▶  Generar Planilla COMISIONES", type="primary", key="btn_com", use_container_width=True):
            with st.spinner("Procesando Comisiones…"):
                rows_com = procesar_comisiones(df_src, int(cons_cp), int(cons_dv))
            if not rows_com:
                st.warning("No se encontraron movimientos de comisiones.")
            else:
                n_cp = sum(1 for r in rows_com if r["Comprobante"] == "CP")
                n_dv = sum(1 for r in rows_com if r["Comprobante"] == "DV")
                st.success(f"✅ **{len(rows_com):,} filas** generadas — CP: **{n_cp}** · DV: **{n_dv}** · Contrapartidas incluidas")
                col_r1, col_r2 = st.columns([1, 2])
                with col_r1:
                    st.markdown("**Resumen por cuenta**")
                    st.dataframe(resumen_cuentas(rows_com), use_container_width=True, hide_index=True, height=320)
                with col_r2:
                    with st.expander("Vista previa — primeras 60 filas", expanded=True):
                        cols_show = [c for c in pd.DataFrame(rows_com).columns if not c.startswith("_")]
                        st.dataframe(pd.DataFrame(rows_com)[cols_show].head(60), use_container_width=True)
                st.markdown("---")
                buf = generar_excel(rows_com, "COMISIONES_RM2")
                st.download_button("⬇️  Descargar Plantilla Comisiones RM2.xlsx", data=buf,
                                   file_name="Plantilla_Comisiones_RM2.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True, type="primary", key="dl_com")
