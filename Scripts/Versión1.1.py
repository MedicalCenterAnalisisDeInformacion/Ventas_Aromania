import json
import calendar
import unicodedata
import base64
from pathlib import Path
import pandas as pd
import numpy as np
from datetime import date, timedelta
# Ajustar rutas y el párametro
EXCEL_PATH  = r"C:/Users/adelarosa/Documents/Reportes/Dashboards/DashboardVentasDiarias_Aromania/08_Agosto/21-08-2026/Dataset.xlsx"
OUTPUT_PATH = r"C:/Users/adelarosa/Documents/Reportes/Dashboards/DashboardVentasDiarias_Aromania/08_Agosto/21-08-2026/index.html"
BOL_EXCLUIR = ["BOLEUCH", "BOLEUGDE", "BOLEUMIN"]
FECHA_BASE  = date(2026, 8, 21)
# Logo del header (opcional). Debe ser un PNG con fondo transparente. Si el
# archivo no existe en esta ruta, el logo simplemente se omite (el header se
# ve igual que antes, sin logo) y el script no se interrumpe.
LOGO_PATH = r"C:/Users/adelarosa/Documents/Reportes/Dashboards/DashboardVentasDiarias_Aromania/Logos/logo.png"
# Lógica de procesamiento
MESES_ES = ["enero","febrero","marzo","abril","mayo","junio",
            "julio","agosto","septiembre","octubre","noviembre","diciembre"]
DIAS_ES  = ["lunes","martes","miércoles","jueves","viernes","sábado","domingo"]
MES_A_NUM = {
    "ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,
    "jul":7,"ago":8,"sep":9,"oct":10,"nov":11,"dic":12,
    "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
    "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12
}
def _margen_seguro(utilidad, ventas):
    m = (utilidad / ventas).replace([np.inf, -np.inf], 0).fillna(0)
    return m.round(4)
def _mapear_mes(serie_mes, origen: str):
    serie_norm = serie_mes.str.strip().str.lower()
    mes_num = serie_norm.map(MES_A_NUM)
    no_reconocidos = serie_norm[mes_num.isna()]
    if len(no_reconocidos):
        valores = sorted(no_reconocidos.unique().tolist())
        print(f"⚠️  [{origen}] {len(no_reconocidos)} fila(s) con valor de 'Mes' no reconocido "
              f"(se agruparán como '?'): {valores}")
    return mes_num
def _formatear_periodo(mes_num_serie, anio_serie):
    mes_abr  = mes_num_serie.apply(lambda m: MESES_ES[int(m)-1].capitalize()[:3] if m and m > 0 else "?")
    anio_abr = anio_serie.astype(int).astype(str).str[-2:]
    return mes_abr + "-" + anio_abr
def _normalizar_texto(s) -> str:
    s = " ".join(str(s).strip().split())
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return s.casefold()
def _detectar_columna(df, patron: str):
    patron_norm = _normalizar_texto(patron)
    for c in df.columns:
        if patron_norm in _normalizar_texto(c):
            return c
    return None
def _cargar_imagen_b64(path: str) -> str:
    """Carga una imagen (p.ej. el logo del header) y la codifica en base64
    para incrustarla directamente en el HTML, sin depender de un archivo
    externo que viaje junto al reporte. Si el archivo no existe, retorna
    cadena vacía y el <img> del header simplemente no se muestra (no rompe
    la generación del dashboard)."""
    p = Path(path)
    if not p.exists():
        print(f"⚠️  No se encontró el logo en {path}; se omite del header.")
        return ""
    return base64.b64encode(p.read_bytes()).decode("utf-8")
def formatear_fechas(base: date):
    ayer = base - timedelta(days=1)
    fecha_reporte = f"{base.day} de {MESES_ES[base.month-1].capitalize()} de {base.year}"
    dia_semana    = DIAS_ES[ayer.weekday()].capitalize()
    fecha_info    = f"{dia_semana}, {ayer.day} de {MESES_ES[ayer.month-1]} de {ayer.year}"
    mes_header    = MESES_ES[base.month-1].capitalize()
    return fecha_reporte, fecha_info, mes_header
def periodo_label_actual(base: date) -> str:
    """Devuelve el PeriodoLabel (p.ej. 'Ago-26') correspondiente a FECHA_BASE,
    con el MISMO formato que _formatear_periodo, para poder anclar en el
    front-end las secciones que deben quedar fijas al mes en curso
    (pestaña 'Resumen Mes en Curso') sin depender del segmentador de meses."""
    mes_abr  = MESES_ES[base.month-1].capitalize()[:3]
    anio_abr = str(base.year)[-2:]
    return f"{mes_abr}-{anio_abr}"
def separar_mes_actual_anterior(vmc, fecha_base):
    """A inicios de mes, 'VentasMesCurso' puede traer mezcladas filas del mes
    que ya cerró (p.ej. julio) junto con las del mes nuevo (p.ej. agosto).
    Esta función separa el DataFrame en:
      - vmc_actual: sólo filas cuyo año/mes coincide con FECHA_BASE (mes en curso real)
      - meses_anteriores: lista de dicts (uno por cada mes "sobrante" detectado)
        con la info necesaria para armar una pestaña "Resumen_<Mes>" congelada.
    Si no hay mezcla (todo el mes coincide con fecha_base), meses_anteriores
    queda vacío y el comportamiento es idéntico al de antes.
    """
    fechas = pd.to_datetime(vmc["Fechas de emisión"])
    es_actual = (fechas.dt.year == fecha_base.year) & (fechas.dt.month == fecha_base.month)
    vmc_actual   = vmc[es_actual].copy()
    vmc_anterior = vmc[~es_actual].copy()
    meses_anteriores = []
    if len(vmc_anterior):
        fechas_ant = pd.to_datetime(vmc_anterior["Fechas de emisión"])
        periodos = sorted(set(zip(fechas_ant.dt.year, fechas_ant.dt.month)))
        for anio, mes in periodos:
            label = MESES_ES[mes-1].capitalize()
            slug  = _normalizar_texto(label).replace(" ", "_")
            mask_periodo = (fechas_ant.dt.year == anio) & (fechas_ant.dt.month == mes)
            df_periodo = vmc_anterior[mask_periodo].copy()
            meses_anteriores.append({
                "anio": anio, "mes": mes, "label": label, "slug": slug, "df": df_periodo
            })
            print(f"ℹ️  [MesCurso] Se detectaron {len(df_periodo)} fila(s) de {label} {anio} mezcladas "
                  f"en 'VentasMesCurso'; se reportarán aparte en la pestaña 'Resumen {label}' "
                  f"y se excluyen del cálculo del mes en curso ({MESES_ES[fecha_base.month-1].capitalize()} "
                  f"{fecha_base.year}).")
    return vmc_actual, meses_anteriores
def procesar_mes_curso(vmc, suc, bol_list):
    vmc["ClaveSucursal"] = pd.to_numeric(vmc["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vmc = vmc.merge(suc, on="ClaveSucursal", how="left")
    vmc["FechaStr"] = pd.to_datetime(vmc["Fechas de emisión"]).dt.strftime("%Y-%m-%d")
    vmc_v = vmc[~vmc["Artículo"].isin(bol_list)]
    agg_v = vmc_v.groupby(["FechaStr","NombreSucursal"]).agg(
        unidades  =("Unidades","sum"),
        ventas    =("Importe c/Desc","sum"),
        utilidad  =("Utilidad","sum"),
    ).reset_index()
    agg_t = vmc.groupby(["FechaStr","NombreSucursal"]).agg(
        tickets=("Movimiento","nunique")
    ).reset_index()
    agg = agg_v.merge(agg_t, on=["FechaStr","NombreSucursal"], how="left")
    agg["tickets"] = agg["tickets"].fillna(0).astype(int)
    agg["margen"]  = _margen_seguro(agg["utilidad"], agg["ventas"])
    for c in ["ventas","utilidad"]:
        agg[c] = agg[c].round(2)
    return agg
def procesar_lineas(vm, art, suc, bol_list):
    vm = vm.copy()
    vm["ClaveSucursal"] = pd.to_numeric(vm["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vm = vm[~vm["Artículo"].isin(bol_list)]
    art_clean = art.drop_duplicates(subset=["Artículo"])
    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])
    vm_merged = vm.merge(art_clean, on="Artículo", how="left").merge(suc_clean, on="ClaveSucursal", how="left")
    vm_merged["Línea"]          = vm_merged["Línea"].fillna("NO ASIGNADO")
    vm_merged["NombreSucursal"] = vm_merged["NombreSucursal"].fillna("OTRO")
    vm_merged["MesNum"] = _mapear_mes(vm_merged["Mes"], "VentasMensuales (Resumen por Línea)").fillna(0).astype(int)
    vm_merged["Año"]    = pd.to_numeric(vm_merged["Año"], errors="coerce").fillna(0).astype(int)
    vm_merged["PeriodoLabel"] = _formatear_periodo(vm_merged["MesNum"], vm_merged["Año"])
    linea_agg = vm_merged.groupby(["NombreSucursal","Línea","PeriodoLabel"]).agg(
        ventas    =("Importe c/Desc","sum"),
        utilidad  =("Utilidad","sum"),
        costo     =("Costo","sum"),
        unidades  =("Unidades","sum"),
    ).reset_index()
    for c in ["ventas","utilidad","costo"]:
        linea_agg[c] = linea_agg[c].round(2)
    linea_agg["margen"] = _margen_seguro(linea_agg["utilidad"], linea_agg["ventas"])
    return linea_agg
def procesar_historico(vm, tkt, suc, bol_list):
    vm["ClaveSucursal"]  = pd.to_numeric(vm["ClaveSucursal"],  errors="coerce").fillna(0).astype(int)
    tkt["ClaveSucursal"] = pd.to_numeric(tkt["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vm = vm[~vm["Artículo"].isin(bol_list)].copy()
    vm["MesNum"] = _mapear_mes(vm["Mes"], "VentasMensuales")
    agg_v = vm.groupby(["Año","MesNum","ClaveSucursal"]).agg(
        unidades  =("Unidades","sum"),
        ventas    =("Importe c/Desc","sum"),
        utilidad  =("Utilidad","sum"),
    ).reset_index()
    tkt["MesNum"] = _mapear_mes(tkt["Mes"], "TicketsMensuales")
    agg_t = tkt.groupby(["Año","MesNum","ClaveSucursal"]).agg(
        tickets=("Tickets","sum")
    ).reset_index()
    agg = agg_v.merge(agg_t, on=["Año","MesNum","ClaveSucursal"], how="outer")
    for c in ["unidades","ventas","utilidad"]:
        agg[c] = agg[c].fillna(0).round(2)
    agg["tickets"] = agg["tickets"].fillna(0).astype(int)
    agg["Año"]    = agg["Año"].fillna(0).astype(int)
    agg["MesNum"] = agg["MesNum"].fillna(0).astype(int)
    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])
    agg = agg.merge(suc_clean, on="ClaveSucursal", how="left")
    agg["NombreSucursal"] = agg["NombreSucursal"].fillna("OTRO")
    agg["margen"] = _margen_seguro(agg["utilidad"], agg["ventas"])
    agg["PeriodoLabel"] = _formatear_periodo(agg["MesNum"], agg["Año"])
    agg = agg.sort_values(["Año","MesNum"]).reset_index(drop=True)
    return agg
def procesar_top_articulos(vm, art_dim, suc, bol_list):
    vm = vm.copy()
    vm["ClaveSucursal"] = pd.to_numeric(vm["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vm = vm[~vm["Artículo"].isin(bol_list)]
    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])
    vm = vm.merge(suc_clean, on="ClaveSucursal", how="left")
    vm["NombreSucursal"] = vm["NombreSucursal"].fillna("OTRO")
    art_clean = art_dim.drop_duplicates(subset=["Artículo"])
    vm = vm.merge(art_clean[["Artículo","Descripción","Fabricante"]], on="Artículo", how="left")
    vm["Descripción"] = vm["Descripción"].fillna("SIN DESCRIPCIÓN")
    vm["Fabricante"]  = vm["Fabricante"].fillna("SIN FABRICANTE")
    vm["MesNum"] = _mapear_mes(vm["Mes"], "VentasMensuales (Top Artículos)").fillna(0).astype(int)
    vm["Año"]    = pd.to_numeric(vm["Año"], errors="coerce").fillna(0).astype(int)
    vm["PeriodoLabel"] = _formatear_periodo(vm["MesNum"], vm["Año"])
    agg = vm.groupby(["Artículo","Descripción","Fabricante","NombreSucursal","PeriodoLabel"]).agg(
        unidades =("Unidades","sum"),
        ventas   =("Importe c/Desc","sum"),
        utilidad =("Utilidad","sum"),
    ).reset_index()
    for c in ["ventas","utilidad"]:
        agg[c] = agg[c].round(2)
    return agg
def procesar_lineas_categoria(vm, art_dim, suc, bol_list):
    vm = vm.copy()
    vm["ClaveSucursal"] = pd.to_numeric(vm["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vm = vm[~vm["Artículo"].isin(bol_list)]
    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])
    vm = vm.merge(suc_clean, on="ClaveSucursal", how="left")
    vm["NombreSucursal"] = vm["NombreSucursal"].fillna("OTRO")
    art_clean = art_dim.drop_duplicates(subset=["Artículo"])
    vm = vm.merge(art_clean[["Artículo","Línea","Categoría"]], on="Artículo", how="left")
    vm["Línea"]     = vm["Línea"].fillna("NO ASIGNADO")
    vm["Categoría"] = vm["Categoría"].fillna("SIN CATEGORÍA")
    vm["MesNum"] = _mapear_mes(vm["Mes"], "VentasMensuales (Líneas y Categorías)").fillna(0).astype(int)
    vm["Año"]    = pd.to_numeric(vm["Año"], errors="coerce").fillna(0).astype(int)
    vm["PeriodoLabel"] = _formatear_periodo(vm["MesNum"], vm["Año"])
    agg = vm.groupby(["Línea","Categoría","NombreSucursal","PeriodoLabel"]).agg(
        unidades =("Unidades","sum"),
        ventas   =("Importe c/Desc","sum"),
        utilidad =("Utilidad","sum"),
    ).reset_index()
    for c in ["ventas","utilidad"]:
        agg[c] = agg[c].round(2)
    return agg
def procesar_categorias_diario(vmc, art_dim, suc, bol_list):
    """Ventas diarias por Línea y Categoría, usado para graficar la tendencia
    día a día en la pestaña 'Resumen Mes Actual'. Se basa en 'VentasMesCurso'
    (ya filtrada al mes real, vmc_actual) cruzada con la dimensión de
    artículos para obtener Línea/Categoría."""
    vmc = vmc.copy()
    vmc["ClaveSucursal"] = pd.to_numeric(vmc["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vmc = vmc[~vmc["Artículo"].isin(bol_list)]
    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])
    vmc = vmc.merge(suc_clean, on="ClaveSucursal", how="left")
    vmc["NombreSucursal"] = vmc["NombreSucursal"].fillna("OTRO")
    art_clean = art_dim.drop_duplicates(subset=["Artículo"])
    vmc = vmc.merge(art_clean[["Artículo","Línea","Categoría"]], on="Artículo", how="left")
    vmc["Línea"]     = vmc["Línea"].fillna("NO ASIGNADO")
    vmc["Categoría"] = vmc["Categoría"].fillna("SIN CATEGORÍA")
    vmc["FechaStr"] = pd.to_datetime(vmc["Fechas de emisión"]).dt.strftime("%Y-%m-%d")
    agg = vmc.groupby(["FechaStr","Línea","Categoría","NombreSucursal"]).agg(
        unidades =("Unidades","sum"),
        ventas   =("Importe c/Desc","sum"),
        utilidad =("Utilidad","sum"),
    ).reset_index()
    for c in ["ventas","utilidad"]:
        agg[c] = agg[c].round(2)
    return agg
def procesar_fabricantes(vm, art_dim, suc, bol_list):
    vm = vm.copy()
    vm["ClaveSucursal"] = pd.to_numeric(vm["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vm = vm[~vm["Artículo"].isin(bol_list)]
    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])
    vm = vm.merge(suc_clean, on="ClaveSucursal", how="left")
    vm["NombreSucursal"] = vm["NombreSucursal"].fillna("OTRO")
    art_clean = art_dim.drop_duplicates(subset=["Artículo"])
    vm = vm.merge(art_clean[["Artículo","Fabricante"]], on="Artículo", how="left")
    vm["Fabricante"] = vm["Fabricante"].fillna("SIN FABRICANTE")
    vm["MesNum"] = _mapear_mes(vm["Mes"], "VentasMensuales (Fabricantes)").fillna(0).astype(int)
    vm["Año"]    = pd.to_numeric(vm["Año"], errors="coerce").fillna(0).astype(int)
    vm["PeriodoLabel"] = _formatear_periodo(vm["MesNum"], vm["Año"])
    agg = vm.groupby(["Fabricante","NombreSucursal","PeriodoLabel"]).agg(
        unidades =("Unidades","sum"),
        ventas   =("Importe c/Desc","sum"),
        utilidad =("Utilidad","sum"),
    ).reset_index()
    for c in ["ventas","utilidad"]:
        agg[c] = agg[c].round(2)
    return agg
def procesar_presupuesto(agg, objetivos, suc, fecha_base):
    dias_mes = calendar.monthrange(fecha_base.year, fecha_base.month)[1]
    fecha_max_global = pd.Timestamp(fecha_base) - pd.Timedelta(days=1)
    primer_dia_mes = pd.Timestamp(year=fecha_base.year, month=fecha_base.month, day=1)
    fin_mes = pd.Timestamp(year=fecha_base.year, month=fecha_base.month, day=dias_mes)
    resumen_ventas = agg.groupby("NombreSucursal").agg(
        unidadesActual=("unidades", "sum"),
        ventasActual=("ventas", "sum"),
        utilidadActual=("utilidad", "sum"),
        fechaMin=("FechaStr", "min"),
    ).reset_index()
    resumen_ventas["fechaMinDt"] = pd.to_datetime(resumen_ventas["fechaMin"])
    # Base: TODAS las sucursales de la dimensión (para incluir también las que
    # aún no tienen venta este mes, con pronóstico 0). Se conserva ClaveSucursal
    # porque 'ObjetivosVentas' puede traer la clave numérica en vez del nombre,
    # y porque se usará para ordenar la tabla final por número de sucursal.
    suc_clean = suc.drop_duplicates(subset=["NombreSucursal"])[["ClaveSucursal", "NombreSucursal"]].copy()
    suc_clean["NombreSucursal"] = suc_clean["NombreSucursal"].astype(str).str.strip()
    resumen_ventas["NombreSucursal"] = resumen_ventas["NombreSucursal"].astype(str).str.strip()
    resumen = suc_clean.merge(resumen_ventas, on="NombreSucursal", how="left")
    for c in ["unidadesActual", "ventasActual", "utilidadActual"]:
        resumen[c] = resumen[c].fillna(0)
    resumen["margen"] = _margen_seguro(resumen["utilidadActual"], resumen["ventasActual"])
    # ── Presupuesto / Objetivo de ventas / Fecha de apertura ──
    col_suc = _detectar_columna(objetivos, "sucursal")
    col_pre = _detectar_columna(objetivos, "presupuesto")
    col_ape = _detectar_columna(objetivos, "apertura")
    resumen["FechaApertura"] = pd.NaT
    if col_suc is None or col_pre is None:
        print("⚠️  [Presupuesto] No se encontraron las columnas 'Sucursal' y/o 'Presupuesto' "
              f"en la hoja 'ObjetivosVentas' (columnas disponibles: {list(objetivos.columns)}). "
              "No se generará la tabla de pronóstico vs. presupuesto.")
        resumen["presupuesto"] = 0.0
    else:
        cols_leer = [col_suc, col_pre] + ([col_ape] if col_ape else [])
        rename_map = {col_suc: "SucursalObjetivo", col_pre: "presupuesto"}
        if col_ape:
            rename_map[col_ape] = "FechaApertura"
        obj_clean = objetivos[cols_leer].rename(columns=rename_map).dropna(subset=["SucursalObjetivo"]).copy()
        obj_clean["presupuesto"] = pd.to_numeric(obj_clean["presupuesto"], errors="coerce").fillna(0)
        if col_ape:
            obj_clean["FechaApertura"] = pd.to_datetime(obj_clean["FechaApertura"], dayfirst=True, errors="coerce")
        else:
            obj_clean["FechaApertura"] = pd.NaT
            print("⚠️  [Presupuesto] No se encontró columna 'FechaApertura' en 'ObjetivosVentas'; "
                  "se usará el método de inferencia anterior (primera venta del mes) para todas las sucursales.")
        claves_num = pd.to_numeric(obj_clean["SucursalObjetivo"], errors="coerce")
        usa_clave = claves_num.notna().mean() >= 0.5 if len(obj_clean) else False
        if usa_clave:
            obj_clean["ClaveSucursal"] = claves_num.fillna(0).astype(int)
            obj_clean = obj_clean.drop_duplicates(subset=["ClaveSucursal"])
            resumen = resumen.drop(columns=["FechaApertura"]).merge(
                obj_clean[["ClaveSucursal", "presupuesto", "FechaApertura"]], on="ClaveSucursal", how="left"
            )
            keys_dim  = set(resumen["ClaveSucursal"])
            keys_obj  = set(obj_clean["ClaveSucursal"])
            huerfanos = sorted(keys_obj - keys_dim)
            modo = "ClaveSucursal (código numérico)"
        else:
            resumen["_key"]     = resumen["NombreSucursal"].apply(_normalizar_texto)
            obj_clean["_key"]   = obj_clean["SucursalObjetivo"].apply(_normalizar_texto)
            obj_clean = obj_clean.drop_duplicates(subset=["_key"])
            resumen = resumen.drop(columns=["FechaApertura"]).merge(
                obj_clean[["_key", "presupuesto", "FechaApertura"]], on="_key", how="left"
            )
            keys_dim = set(resumen["_key"])
            keys_obj = set(obj_clean["_key"])
            huerfanos = sorted(
                obj_clean.loc[obj_clean["_key"].isin(keys_obj - keys_dim), "SucursalObjetivo"].unique().tolist()
            )
            resumen = resumen.drop(columns=["_key"])
            modo = "nombre de sucursal (texto normalizado)"
        if huerfanos:
            print(f"⚠️  [Presupuesto] Cruce por {modo}. {len(huerfanos)} valor(es) en 'ObjetivosVentas' NO "
                  f"encontrados en la dimensión de sucursales: {huerfanos}")
        resumen["presupuesto"] = resumen["presupuesto"].fillna(0)
    antes = len(resumen)
    sin_presupuesto_mask = resumen["presupuesto"] <= 0
    sin_nada_mask = sin_presupuesto_mask & (resumen["ventasActual"] <= 0)
    sin_presupuesto_con_venta = sorted(
        resumen.loc[sin_presupuesto_mask & ~sin_nada_mask, "NombreSucursal"].tolist()
    )
    excluidas = sorted(resumen.loc[sin_nada_mask, "NombreSucursal"].tolist())
    resumen = resumen[~sin_nada_mask].copy()
    resumen.loc[resumen["presupuesto"] <= 0, "presupuesto"] = np.nan
    if sin_presupuesto_con_venta:
        print(f"ℹ️  [Presupuesto] {len(sin_presupuesto_con_venta)} sucursal(es) con venta pero SIN "
              f"presupuesto asignado: se incluyen con presupuesto/cumplimiento en blanco: "
              f"{sin_presupuesto_con_venta}")
    if excluidas:
        print(f"ℹ️  [Presupuesto] {len(excluidas)} de {antes} sucursal(es) se excluyeron por no tener "
              f"presupuesto asignado ni venta en el mes en curso: {excluidas}")
    usa_inferencia = sorted(
        resumen.loc[resumen["FechaApertura"].isna(), "NombreSucursal"].tolist()
    )
    if usa_inferencia:
        print(f"ℹ️  [Presupuesto] {len(usa_inferencia)} sucursal(es) sin 'FechaApertura' registrada: se usó "
              f"el método de inferencia anterior (primera venta del mes) para calcular sus días: {usa_inferencia}")
    def _calcular_dias(row):
        apertura = row["FechaApertura"]
        if pd.isna(apertura):
            inicio = row["fechaMinDt"] if pd.notna(row["fechaMinDt"]) else primer_dia_mes
        else:
            inicio = apertura
        if inicio < primer_dia_mes:
            inicio = primer_dia_mes
        dias_operativos = max(0, (fin_mes - inicio).days + 1)
        dias_transcurridos = max(0, (fecha_max_global - inicio).days + 1)
        return pd.Series({"diasOperativosMes": dias_operativos, "diasTranscurridos": dias_transcurridos})
    resumen[["diasOperativosMes", "diasTranscurridos"]] = resumen.apply(_calcular_dias, axis=1)
    resumen["pronostico"] = np.where(
        resumen["diasTranscurridos"] > 0,
        resumen["ventasActual"] / resumen["diasTranscurridos"] * resumen["diasOperativosMes"],
        0.0,
    )
    resumen["cumplimiento"] = np.where(
        resumen["presupuesto"] > 0,
        _margen_seguro(resumen["pronostico"], resumen["presupuesto"]),
        np.nan,
    )
    for c in ["unidadesActual", "ventasActual", "utilidadActual", "pronostico", "presupuesto"]:
        resumen[c] = resumen[c].round(2)
    resumen = resumen.sort_values("ClaveSucursal").reset_index(drop=True)
    def _nan_a_none(serie):
        return serie.astype(object).where(serie.notna(), None)
    resumen["presupuesto"]  = _nan_a_none(resumen["presupuesto"])
    resumen["cumplimiento"] = _nan_a_none(resumen["cumplimiento"])
    return resumen[["ClaveSucursal", "NombreSucursal", "unidadesActual", "ventasActual", "utilidadActual",
                     "margen", "presupuesto", "cumplimiento", "pronostico"]]
def ordenar_sucursales_por_apertura(suc: pd.DataFrame, objetivos: pd.DataFrame) -> list:
    """Determina el orden de los botones del segmentador de sucursales:
    de la FechaApertura más antigua a la más reciente (columna 'FechaApertura'
    en la hoja 'ObjetivosVentas', la misma que usa procesar_presupuesto para
    el pronóstico). Las sucursales sin FechaApertura registrada se colocan al
    final, ordenadas por ClaveSucursal (ID) ascendente.
    Es una función independiente y autocontenida (no comparte estado con
    procesar_presupuesto) para no arriesgar esa lógica ya probada.
    """
    suc_clean = suc.drop_duplicates(subset=["NombreSucursal"])[["ClaveSucursal", "NombreSucursal"]].copy()
    suc_clean["NombreSucursal"] = suc_clean["NombreSucursal"].astype(str).str.strip()
    suc_clean["FechaApertura"] = pd.NaT
    col_suc = _detectar_columna(objetivos, "sucursal")
    col_ape = _detectar_columna(objetivos, "apertura")
    if col_suc is None or col_ape is None:
        print("⚠️  [Segmentador Sucursales] No se encontró columna 'Sucursal' y/o 'FechaApertura' "
              "en 'ObjetivosVentas'; las sucursales se ordenarán únicamente por ClaveSucursal (ID).")
    else:
        obj_clean = objetivos[[col_suc, col_ape]].rename(
            columns={col_suc: "SucursalObjetivo", col_ape: "FechaApertura"}
        ).dropna(subset=["SucursalObjetivo"]).copy()
        obj_clean["FechaApertura"] = pd.to_datetime(obj_clean["FechaApertura"], dayfirst=True, errors="coerce")
        claves_num = pd.to_numeric(obj_clean["SucursalObjetivo"], errors="coerce")
        usa_clave = claves_num.notna().mean() >= 0.5 if len(obj_clean) else False
        if usa_clave:
            obj_clean["ClaveSucursal"] = claves_num.fillna(0).astype(int)
            obj_clean = obj_clean.drop_duplicates(subset=["ClaveSucursal"])
            suc_clean = suc_clean.drop(columns=["FechaApertura"]).merge(
                obj_clean[["ClaveSucursal", "FechaApertura"]], on="ClaveSucursal", how="left"
            )
        else:
            suc_clean["_key"] = suc_clean["NombreSucursal"].apply(_normalizar_texto)
            obj_clean["_key"] = obj_clean["SucursalObjetivo"].apply(_normalizar_texto)
            obj_clean = obj_clean.drop_duplicates(subset=["_key"])
            suc_clean = suc_clean.drop(columns=["FechaApertura"]).merge(
                obj_clean[["_key", "FechaApertura"]], on="_key", how="left"
            )
            suc_clean = suc_clean.drop(columns=["_key"])
    sin_fecha_n = suc_clean["FechaApertura"].isna().sum()
    if sin_fecha_n:
        nombres_sin_fecha = sorted(suc_clean.loc[suc_clean["FechaApertura"].isna(), "NombreSucursal"].tolist())
        print(f"ℹ️  [Segmentador Sucursales] {sin_fecha_n} sucursal(es) sin FechaApertura registrada; "
              f"se colocarán al final, ordenadas por ClaveSucursal: {nombres_sin_fecha}")
    con_fecha = suc_clean[suc_clean["FechaApertura"].notna()].sort_values("FechaApertura")
    sin_fecha = suc_clean[suc_clean["FechaApertura"].isna()].sort_values("ClaveSucursal")
    orden_final = pd.concat([con_fecha, sin_fecha], ignore_index=True)
    return orden_final["NombreSucursal"].tolist()
def generar_html(agg, linea_agg, historico_agg, top_art_agg, lineas_cat_agg, fabricantes_agg,
                 presupuesto_agg, categorias_diario_agg, lista_sucursales, fecha_reporte, fecha_info, mes_header,
                 current_period_label, resumenes_anteriores=None):
    resumenes_anteriores = resumenes_anteriores or []
    # ── Pestañas extra "Resumen <Mes>" (mes(es) cerrados detectados mezclados
    # en VentasMesCurso a inicios de mes). Si no hay ninguno, estas variables
    # quedan vacías y el dashboard se ve exactamente igual que antes. ──
    tabs_nav_extra = ""
    tabs_content_extra = ""
    resumenes_ant_dict = {}
    for r in resumenes_anteriores:
        slug, label, anio = r["slug"], r["label"], r["anio"]
        tabid = f"resumen_{slug}"
        tabs_nav_extra += (f'\n  <button class="tab-nav-btn" id="tabnav-{tabid}" '
                            f'onclick="switchTab(\'{tabid}\')">Resumen {label}</button>')
        tabs_content_extra += f"""
<div id="tab-{tabid}" class="tab-content">
<div class="main">
  <div class="tc">
    <div class="card-head">
      <div><div class="card-title">Ventas Diarias · {label} {anio}</div><div class="card-sub">Mes cerrado · detalle completo por día</div></div>
      <span class="note-bol">Sucursales según selección</span>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead><tr><th>Fecha</th><th>Día</th><th class="r">Unidades</th><th class="r">Ventas $</th><th class="r">Utilidad</th><th class="r">Margen</th><th class="r">Tickets</th></tr></thead>
      <tbody id="tabla-resumen-{slug}"></tbody>
    </table>
    </div>
  </div>
  <div class="charts-row">
    <div class="cc"><div class="card-head" style="margin-bottom:.4rem"><div><div class="card-title">Ventas $</div><div class="card-sub">Volumen diario · Sucursales seleccionadas</div></div></div><div class="cw"><canvas id="chart-resumen-{slug}-ventas"></canvas></div></div>
    <div class="cc"><div class="card-head" style="margin-bottom:.4rem"><div><div class="card-title">No. de Tickets</div><div class="card-sub">Volumen diario · Sucursales seleccionadas</div></div></div><div class="cw"><canvas id="chart-resumen-{slug}-tickets"></canvas></div></div>
  </div>
</div>
</div>"""
        resumenes_ant_dict[slug] = {"label": label, "anio": anio, "data": r["agg"].to_dict("records")}
    resumenes_ant_json = json.dumps(resumenes_ant_dict, ensure_ascii=False)
    data_json         = json.dumps(agg.to_dict("records"),           ensure_ascii=False)
    linea_json        = json.dumps(linea_agg.to_dict("records"),     ensure_ascii=False)
    historico_json    = json.dumps(historico_agg.to_dict("records"), ensure_ascii=False)
    top_art_json      = json.dumps(top_art_agg.to_dict("records"),   ensure_ascii=False)
    lineas_cat_json   = json.dumps(lineas_cat_agg.to_dict("records"),ensure_ascii=False)
    fabricantes_json  = json.dumps(fabricantes_agg.to_dict("records"),ensure_ascii=False)
    presupuesto_json  = json.dumps(presupuesto_agg.to_dict("records"),ensure_ascii=False)
    categorias_diario_json = json.dumps(categorias_diario_agg.to_dict("records"), ensure_ascii=False)
    sucursales_json   = json.dumps(lista_sucursales,                  ensure_ascii=False)
    current_period_json = json.dumps(current_period_label,            ensure_ascii=False)
    logo_b64 = _cargar_imagen_b64(LOGO_PATH)
    periodos_unicos = (
        historico_agg[["Año","MesNum","PeriodoLabel"]]
        .drop_duplicates()
        .sort_values(["Año","MesNum"])["PeriodoLabel"]
        .tolist()
    )
    periodos_json = json.dumps(periodos_unicos, ensure_ascii=False)
    html_template = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Aromania · Análisis de Ventas</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-chart-treemap@3/dist/chartjs-chart-treemap.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
html{font-size:clamp(13px, 4vw, 16px)}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#EFF8FC;color:#16232B;min-height:100vh}
header{background:#00B0F0;padding:1.2rem 2rem;display:grid;grid-template-columns:1fr auto 1fr;column-gap:1rem;align-items:center;box-shadow:0 3px 20px rgba(0,176,240,0.2)}
.header-left{min-width:0}
.header-left h1{font-size:1.15rem;font-weight:700;color:#fff;letter-spacing:.03em}
.header-left p{font-size:.72rem;color:#EAF8FF;margin-top:4px;letter-spacing:0.02em}
.header-right{display:flex;flex-direction:column;align-items:flex-end;gap:.35rem;min-width:0}
.header-logo{display:flex;align-items:center;justify-content:center;min-width:0}
.header-logo img{height:5.4rem;width:auto;display:block}
.hbadge{background:#ffffff1f;border:1px solid #ffffff40;color:#fff;font-size:.68rem;font-weight:600;padding:4px 14px;border-radius:20px;letter-spacing:.05em}
.hdate{font-size:.72rem;color:#EAF8FF;font-weight:500}
.filter-bar{background:#fff;border-bottom:1px solid #D9EEF7;padding:.75rem 2rem;display:flex;gap:.8rem;flex-wrap:wrap;align-items:center;justify-content:center;box-shadow:0 2px 4px rgba(0,176,240,0.02)}
.filter-bar + .filter-bar{border-top:1px solid #EAF7FC;box-shadow:none}
.suc-label{font-size:.67rem;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:#00B0F0;margin-right:.25rem;white-space:nowrap}
.suc-btns{display:flex;gap:6px;flex-wrap:wrap;justify-content:center}
.suc-btn{display:inline-flex;align-items:center;padding:6px 13px;border-radius:20px;border:1px solid #D2ECF5;background:#E3F5FC;color:#46626D;font-size:.72rem;font-weight:600;cursor:pointer;transition:all .2s;font-family:inherit;white-space:nowrap}
.suc-btn.active{background:#00B0F0 !important;border-color:#00B0F0 !important;color:#fff !important;font-weight:700;box-shadow:0 3px 10px rgba(0,176,240,0.2)}
.mes-btn{display:inline-flex;align-items:center;padding:5px 11px;border-radius:20px;border:1px solid #D2ECF5;background:#E3F5FC;color:#46626D;font-size:.69rem;font-weight:600;cursor:pointer;transition:all .2s;font-family:inherit;white-space:nowrap}
.mes-btn.active{background:#14323D !important;border-color:#14323D !important;color:#fff !important;font-weight:700}
.suc-btn.oculto,.mes-btn.oculto{display:none}
.suc-sep{width:1px;height:20px;background:#CDEAF8;margin:0 .1rem;flex-shrink:0}
.btn-all{padding:6px 13px;border-radius:20px;border:1px solid #00B0F0;background:#fff;color:#00B0F0;font-size:.7rem;font-weight:700;cursor:pointer;font-family:inherit;transition:all .2s;white-space:nowrap}
.btn-all.dark{border-color:#14323D;color:#14323D}
.tabs-nav{background:#fff;border-bottom:2px solid #D2ECF5;padding:0 2rem;display:flex;justify-content:center;gap:0;flex-wrap:wrap}
.tab-nav-btn{padding:.85rem 1.6rem;font-size:.82rem;font-weight:700;color:#666666;cursor:pointer;border:none;background:none;font-family:inherit;border-bottom:3px solid transparent;margin-bottom:-2px;transition:all .2s;letter-spacing:.02em}
.tab-nav-btn.active{color:#00B0F0;border-bottom-color:#00B0F0}
.tab-nav-btn:hover:not(.active){color:#14323D;border-bottom-color:#D2ECF5}
.tab-content{display:none}
.tab-content.active{display:block}
.main{padding:1.5rem 2rem 3rem;max-width:1450px;margin:0 auto}
.kpi-grid{display:flex;flex-wrap:wrap;justify-content:center;gap:.9rem;margin-bottom:1.5rem}
.kpi{background:#fff;border-radius:12px;padding:1.1rem 1.2rem;border:1px solid #D2ECF5;position:relative;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.015);flex:1 1 170px;max-width:215px;min-width:0;text-align:center}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--ac,#00B0F0)}
.kpi-label{font-size:.63rem;font-weight:700;text-transform:uppercase;letter-spacing:.09em;color:#666666;margin-bottom:6px}
.kpi-value{font-size:1.4rem;font-weight:700;color:#16232B;line-height:1}
.kpi-value.purple{color:#00B0F0}
.kpi-value.gold{color:#aa7300}
.kpi-sub{font-size:.67rem;color:#666666;margin-top:5px}
.kpi-note{font-size:.68rem;color:#666666;margin:-.6rem 0 1rem;font-style:italic}
.tc{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #D2ECF5;margin-bottom:.9rem;box-shadow:0 2px 6px rgba(0,0,0,0.015)}
.card-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:1.1rem;gap:1rem;flex-wrap:wrap}
.card-title{font-size:.88rem;font-weight:700;color:#16232B;letter-spacing:0.01em}
.card-sub{font-size:.7rem;color:#666666;margin-top:2px}
.note-bol{font-size:.67rem;background:#E3F5FC;color:#00B0F0;border:1px solid #CDEAF8;padding:4px 10px;border-radius:6px;white-space:nowrap;font-weight:500}
.metric-tabs{display:flex;gap:4px;background:#EAF7FC;padding:3px;border-radius:8px;border:1px solid #D2ECF5}
.tab-btn{background:none;border:none;padding:5px 12px;font-size:.72rem;font-weight:600;color:#46626D;cursor:pointer;border-radius:6px;font-family:inherit;transition:all .15s}
.tab-btn.active{background:#fff;color:#00B0F0;box-shadow:0 2px 5px rgba(0,176,240,0.1);font-weight:700}
table{width:100%;border-collapse:collapse;font-size:.78rem;min-width:100%}
thead tr{background:#F7FCFE;border-bottom:2px solid #D2ECF5}
thead th{text-align:left;padding:9px 11px;font-size:.63rem;font-weight:700;text-transform:uppercase;letter-spacing:.09em;color:#00B0F0;white-space:nowrap}
thead th.r{text-align:right}
tbody tr{border-bottom:1px solid #EAF7FC;transition:background .15s}
tbody tr:last-child{border-bottom:none}
tbody tr:hover{background:#E3F5FC}
tbody tr.total-row{background:linear-gradient(90deg,#EFF8FC,#fff);border-top:2px solid #7FD1EF}
tbody tr.total-row td{font-weight:700;color:#00B0F0}
td{padding:9px 11px;color:#2B3A42}
td.r{text-align:right;font-variant-numeric:tabular-nums}
td.date{font-weight:600;color:#16232B;white-space:nowrap}
td.dayname{font-size:.68rem;color:#666666;white-space:nowrap}
td.art-code{font-family:'Consolas',monospace;font-size:.72rem;color:#00B0F0;font-weight:700;white-space:nowrap}
td.art-desc{font-size:.75rem;color:#2B3A42;line-height:1.4;word-break:break-word;min-width:200px}
td.art-fab{font-size:.72rem;color:#46626D;word-break:break-word;min-width:120px}
td.rank{font-size:.78rem;font-weight:800;color:#666666;width:32px;text-align:center}
td.rank.top3{color:#00B0F0}
.pill{display:inline-block;padding:2px 9px;border-radius:20px;font-size:.65rem;font-weight:600;letter-spacing:0.02em}
.pill.hi{background:#e3f7ed;color:#176440}
.pill.mi{background:#fff3db;color:#805200}
.pill.lo{background:#fbe4e4;color:#a12727}
.charts-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:.9rem}
.cc{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #D2ECF5;box-shadow:0 2px 6px rgba(0,0,0,0.015)}
.lines-container{display:flex;flex-direction:column;gap:.9rem;margin-top:1.5rem}
.lines-table-box,.lines-chart-box{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #D2ECF5;box-shadow:0 2px 6px rgba(0,0,0,0.015);display:flex;flex-direction:column;min-width:0}
.hist-card{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #D2ECF5;box-shadow:0 2px 6px rgba(0,0,0,0.015)}
.cw{position:relative;height:265px}
.cw.bars-horizontal{min-height:220px}
.cw.hist-chart{position:relative;height:360px;width:100%}
.cw.bubble-chart{position:relative;height:560px;width:100%}
.empty{text-align:center;padding:2.5rem;color:#666666;font-size:.82rem}
.section-divider{border:none;border-top:2px solid #ede9f3;margin:2rem 0}
.top-header-card{background:#fff;border-radius:12px;padding:1rem 1.6rem;border:1px solid #D2ECF5;box-shadow:0 2px 6px rgba(0,0,0,0.015);margin-bottom:1rem}
.top-layout{display:block}
.top-table-box{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #D2ECF5;box-shadow:0 2px 6px rgba(0,0,0,0.015);margin-bottom:1rem}
.top-bubble-box{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #D2ECF5;box-shadow:0 2px 6px rgba(0,0,0,0.015);display:flex;flex-direction:column;min-height:0}
.badge-rank{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;font-size:.68rem;font-weight:800;background:#EAF7FC;color:#666666}
.badge-rank.r1{background:#00B0F0;color:#fff}
.badge-rank.r2{background:#1E88B8;color:#fff}
.badge-rank.r3{background:#6FD3F7;color:#fff}
.top-kpi-row{display:grid;grid-template-columns:repeat(3,1fr);gap:.75rem;margin-bottom:1rem}
.top-kpi{background:#fff;border-radius:10px;padding:.9rem 1.1rem;border:1px solid #D2ECF5;position:relative;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.015);text-align:center}
.top-kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--ac,#00B0F0)}
.top-kpi .kpi-label{font-size:.6rem}
.top-kpi .kpi-value{font-size:1.1rem}
.legend-gradient{display:inline-flex;align-items:center;gap:6px;font-size:.67rem;color:#46626D;white-space:nowrap}
.legend-bar{width:70px;height:8px;border-radius:4px;background:linear-gradient(90deg,#CFECF8,#045C7F)}
tr.lc-parent{cursor:pointer}
tr.lc-parent:hover{background:#E3F5FC}
tr.lc-child td{background:#E3F5FC;color:#46626D}
tr.lc-child td:nth-child(2){padding-left:26px}
.lc-icon{display:inline-block;width:14px;color:#00B0F0;font-weight:700}
.fq-badge{display:inline-block;padding:2px 10px;border-radius:20px;font-size:.65rem;font-weight:700;letter-spacing:.02em;white-space:nowrap}
tr.fq-parent{cursor:pointer}
tr.fq-parent:hover td{filter:brightness(0.97)}
tr.fq-parent td{padding-top:10px;padding-bottom:10px}
.fab-bar-cell{position:relative}
.fab-bar-bg{position:absolute;right:0;top:3px;bottom:3px;border-radius:4px;z-index:0}
.fab-bar-cell b{position:relative;z-index:1}
.fab-rank-num{font-size:.68rem;color:#666666;font-weight:700}
.table-scale-wrap{
  width:100%;
  overflow:hidden;
}
.table-scale-wrap table{
  transform-origin: top left;
}
@media (max-width: 640px){
  .main{padding:1.2rem 1rem 2.5rem}
  .kpi-grid{gap:.4rem}
  .kpi{flex:1 1 calc(33.333% - .4rem);max-width:none;padding:.7rem .4rem;min-width:0}
  .kpi:last-child:nth-child(3n+1){flex-basis:100%}
  .kpi:nth-last-child(2):nth-child(3n+1),
  .kpi:last-child:nth-child(3n+2){flex-basis:calc(50% - .4rem)}
  .kpi-value{font-size:clamp(.72rem, 4.1vw, 1.15rem);overflow-wrap:anywhere}
  .kpi-label{font-size:clamp(.42rem, 2.2vw, .58rem);letter-spacing:.03em;line-height:1.2;margin-bottom:4px;overflow-wrap:break-word;hyphens:auto}
  .kpi-sub{font-size:clamp(.4rem, 2vw, .58rem);line-height:1.2;margin-top:3px;overflow-wrap:break-word;hyphens:auto}
  .top-kpi-row{grid-template-columns:repeat(3,1fr) !important;gap:.4rem}
  .top-kpi:last-child:nth-child(3n+1){grid-column:1 / -1}
  .top-kpi:nth-last-child(2):nth-child(3n+1){grid-column:span 2}
  .top-kpi{padding:.65rem .4rem;min-width:0}
  .top-kpi .kpi-value{font-size:clamp(.68rem, 3.4vw, .95rem);overflow-wrap:anywhere}
  .top-kpi .kpi-label{font-size:clamp(.4rem, 2.1vw, .55rem);line-height:1.2}
  .top-kpi .kpi-value.kv-txt{font-size:clamp(.6rem, 2.9vw, .85rem) !important}
  .filter-bar{padding:.5rem .75rem;gap:.4rem}
  .suc-label{font-size:.6rem;margin-right:0}
  .suc-btns{gap:4px}
  .suc-btn,.mes-btn{padding:3px 8px;font-size:.6rem}
  .btn-all{padding:4px 10px;font-size:.6rem}
  .suc-sep{height:14px}
  .tab-nav-btn{padding:.7rem .9rem;font-size:.74rem}
  header{column-gap:.5rem;padding:1rem 1rem}
  .header-logo img{height:3.6rem}
}
</style>
</head>
<body>
<header>
  <div class="header-left">
    <h1>Aromania — Análisis de Ventas</h1>
    <p><strong>Período __MES_HEADER__</strong></p>
    <p>Elaborado con información al día __FECHA_INFO__</p>
  </div>
  <div class="header-logo"><img src="data:image/png;base64,__LOGO_B64__" alt="Aromania"></div>
  <div class="header-right">
    <span class="hdate">Reporte Generado: __FECHA_VALOR__</span>
    <div class="hbadge">DASHBOARD v1.1</div>
  </div>
</header>
<div class="filter-bar" id="filter-bar-suc">
  <span class="suc-label">Sucursales</span>
  <div class="suc-btns" id="suc-btns"></div>
  <div class="suc-sep"></div>
  <button class="btn-all" onclick="toggleAll()">Todas / Ninguna</button>
</div>
<div class="filter-bar" id="filter-bar-mes">
  <span class="suc-label" style="color:#14323D">Meses · Histórico</span>
  <div class="suc-btns" id="mes-btns"></div>
  <div class="suc-sep"></div>
  <button class="btn-all dark" onclick="toggleAllMeses()">Todos / Ninguno</button>
</div>
<div class="tabs-nav">
  <button class="tab-nav-btn active" id="tabnav-objetivos"    onclick="switchTab('objetivos')">Ventas vs. Pronóstico Mes en Curso</button>__TABS_NAV_EXTRA__
  <button class="tab-nav-btn"        id="tabnav-resumenactual"    onclick="switchTab('resumenactual')">Resumen Mes Actual</button>
  <button class="tab-nav-btn"        id="tabnav-resumenacumulado" onclick="switchTab('resumenacumulado')">Resumen Histórico</button>
  <button class="tab-nav-btn"        id="tabnav-toparticulos" onclick="switchTab('toparticulos')">Top 10 Artículos</button>
  <button class="tab-nav-btn"        id="tabnav-lineascategoria" onclick="switchTab('lineascategoria')">Líneas y Categorías</button>
  <button class="tab-nav-btn"        id="tabnav-fabricantes" onclick="switchTab('fabricantes')">Fabricantes</button>
</div>
<div id="tab-objetivos" class="tab-content active">
<div class="main">
  <div class="tc" style="margin-bottom:.9rem">
    <div class="card-head">
      <div><div class="card-title">Ventas vs. Pronóstico</div><div class="card-sub">Ventas $ mes en curso</div></div>
      <span class="note-bol">Pronóstico de Alcance = Ventas estimadas al cierre de mes</span>
    </div>
    <div class="top-kpi-row" id="obj-kpi-row" style="grid-template-columns:repeat(4,1fr)">
      <div class="top-kpi" style="--ac:#14323D"><div class="kpi-label">Ventas $</div><div class="kpi-value" id="p-ventas">—</div><div class="kpi-sub">Acumulado mes en curso</div></div>
      <div class="top-kpi" style="--ac:#00B0F0"><div class="kpi-label">Pronóstico de cierre</div><div class="kpi-value purple" id="p-pronostico">—</div><div class="kpi-sub">Suma de pronósticos</div></div>
      <div class="top-kpi" id="kpi-presupuesto" style="--ac:#aa7300"><div class="kpi-label">Presupuesto</div><div class="kpi-value gold" id="p-presupuesto">—</div><div class="kpi-sub">Acumulado (sólo sucursales con presupuesto asignado)</div></div>
      <div class="top-kpi" id="kpi-cumplimiento" style="--ac:#1A7A4A"><div class="kpi-label">Alcance actual</div><div class="kpi-value" id="p-cumplimiento" style="color:#1A7A4A">—</div><div class="kpi-sub">Ventas actuales ÷ Presupuesto</div></div>
    </div>
    <div class="table-scale-wrap">
      <table>
        <thead><tr><th>Sucursal</th><th class="r">Unidades</th><th class="r">Venta $</th><th class="r">Utilidad</th><th class="r">Margen</th><th class="r" id="th-presupuesto">Presupuesto</th><th class="r" id="th-alcance">Alcance actual</th><th class="r">Pronóstico de Alcance</th></tr></thead>
        <tbody id="tabla-presupuesto"></tbody>
      </table>
    </div>
  </div>
  <div class="hist-card">
    <div class="card-head">
      <div><div class="card-title">Ventas vs. Objetivo</div><div class="card-sub">Comparativo por sucursal</div></div>
    </div>
    <div class="cw" id="obj-chart-wrap" style="height:420px"><canvas id="chart-objetivos"></canvas></div>
  </div>
</div>
</div>
<div id="tab-resumenactual" class="tab-content">
<div class="main">
  <p class="kpi-note">* Esta pestaña sólo muestra información correspondiente al mes en curso (__MES_HEADER__).</p>
  <div class="kpi-grid">
    <div class="kpi" style="--ac:#1E88B8"><div class="kpi-label">Unidades vendidas</div><div class="kpi-value" id="k-uni">—</div><div class="kpi-sub">Ventas mes en curso</div></div>
    <div class="kpi" style="--ac:#00B0F0"><div class="kpi-label">Ventas $</div><div class="kpi-value purple" id="k-ventas">—</div><div class="kpi-sub" id="k-ventas-s">—</div></div>
    <div class="kpi" style="--ac:#045C7F"><div class="kpi-label">Utilidad</div><div class="kpi-value purple" id="k-util">—</div><div class="kpi-sub" id="k-util-sub">Margen: —</div></div>
    <div class="kpi" style="--ac:#14323D"><div class="kpi-label">Tickets</div><div class="kpi-value" id="k-tkt">—</div><div class="kpi-sub">Volumen de ventas</div></div>
    <div class="kpi" style="--ac:#4FA8D8"><div class="kpi-label">Ventas $ promedio por ticket</div><div class="kpi-value" id="k-vtkt">—</div><div class="kpi-sub">Ventas $ ÷ Tickets</div></div>
    <div class="kpi" style="--ac:#7FD1EF"><div class="kpi-label">Unidades promedio por Ticket</div><div class="kpi-value" id="k-utkt">—</div><div class="kpi-sub">Unidades ÷ Tickets</div></div>
  </div>
  <div class="tc">
    <div class="card-head">
      <div><div class="card-title">Resumen de Ventas · Mes en curso</div><div class="card-sub">Detalle diario según sucursales seleccionadas</div></div>
      <span class="note-bol">Desempeño de ventas por día</span>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead><tr><th>Fecha</th><th>Día</th><th class="r">Unidades</th><th class="r">Ventas $</th><th class="r">Utilidad</th><th class="r">Margen</th><th class="r">Tickets</th></tr></thead>
      <tbody id="tabla-body"></tbody>
    </table>
    </div>
  </div>
  <div class="charts-row">
    <div class="cc"><div class="card-head" style="margin-bottom:.4rem"><div><div class="card-title">Ventas $</div><div class="card-sub">Volumen diario · Sucursales seleccionadas</div></div></div><div class="cw"><canvas id="chart-ventas"></canvas></div></div>
    <div class="cc"><div class="card-head" style="margin-bottom:.4rem"><div><div class="card-title">No. de Tickets</div><div class="card-sub">Volumen diario · Sucursales seleccionadas</div></div></div><div class="cw"><canvas id="chart-tickets"></canvas></div></div>
  </div>
  <hr class="section-divider">
  <div class="tc" style="margin-bottom:.9rem">
    <div class="card-head">
      <div><div class="card-title">Comparativo por Sucursal · Mes en Curso</div><div class="card-sub">Unidades, ventas, utilidad y margen acumulados del mes</div></div>
      <span class="note-bol">Sólo mes en curso</span>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead><tr><th>Sucursal</th><th class="r">Unidades</th><th class="r">Ventas $</th><th class="r">Utilidad</th><th class="r">Margen</th><th class="r">Tickets</th></tr></thead>
      <tbody id="tabla-sucursal-actual"></tbody>
    </table>
    </div>
  </div>
  <div class="hist-card">
    <div class="card-head">
      <div><div class="card-title">Comparativo por Sucursal</div><div class="card-sub">Métrica seleccionada + Margen % · Mes en curso</div></div>
      <div class="metric-tabs">
        <button class="tab-btn active" id="btn-hs-ventas"   onclick="changeSucMetric('ventas')">Ventas c/Desc</button>
        <button class="tab-btn"        id="btn-hs-utilidad" onclick="changeSucMetric('utilidad')">Utilidad</button>
        <button class="tab-btn"        id="btn-hs-unidades" onclick="changeSucMetric('unidades')">Unidades</button>
        <button class="tab-btn"        id="btn-hs-tickets"  onclick="changeSucMetric('tickets')">Tickets</button>
      </div>
    </div>
    <div class="cw" id="sucursalactual-chart-wrap" style="height:340px"><canvas id="chart-sucursalactual"></canvas></div>
  </div>
  <hr class="section-divider">
  <div class="lines-container">
    <div class="lines-table-box">
      <div class="card-head">
        <div><div class="card-title">Resumen de Ventas por Línea - Categoría · Mes en Curso</div><div class="card-sub">Sólo mes en curso</div></div>
      </div>
      <div class="table-scale-wrap">
        <table>
          <thead><tr><th style="width:26px"></th><th>Línea</th><th class="r">Unidades</th><th class="r">Ventas $</th><th class="r">Utilidad</th><th class="r">Margen</th></tr></thead>
          <tbody id="tabla-lineas-actual"></tbody>
        </table>
      </div>
    </div>
    <div class="lines-chart-box">
      <div class="card-head" style="margin-bottom:.7rem">
        <div><div class="card-title">Tendencia de Categorías</div><div class="card-sub">Evolución diaria por categoría · Mes en curso</div></div>
        <div class="metric-tabs">
          <button class="tab-btn active" id="btn-ma-ventas"   onclick="changeLineMetricActual('ventas')">Ventas c/Desc</button>
          <button class="tab-btn"        id="btn-ma-unidades" onclick="changeLineMetricActual('unidades')">Unidades</button>
          <button class="tab-btn"        id="btn-ma-utilidad" onclick="changeLineMetricActual('utilidad')">Utilidad</button>
        </div>
      </div>
      <div class="cw" id="lineasactual-chart-wrap"><canvas id="chart-lineasactual"></canvas></div>
    </div>
  </div>
</div>
</div>
<div id="tab-resumenacumulado" class="tab-content">
<div class="main">
  <p class="kpi-note">* Las tarjetas muestran información de acuerdo a lo seleccionado en los segmentadores.</p>
  <div class="kpi-grid">
    <div class="kpi" style="--ac:#1E88B8"><div class="kpi-label">Unidades vendidas</div><div class="kpi-value" id="ka-uni">—</div><div class="kpi-sub">Acumulado según selección</div></div>
    <div class="kpi" style="--ac:#00B0F0"><div class="kpi-label">Ventas $</div><div class="kpi-value purple" id="ka-ventas">—</div><div class="kpi-sub" id="ka-ventas-s">—</div></div>
    <div class="kpi" style="--ac:#045C7F"><div class="kpi-label">Utilidad</div><div class="kpi-value purple" id="ka-util">—</div><div class="kpi-sub" id="ka-util-sub">Margen: —</div></div>
    <div class="kpi" style="--ac:#14323D"><div class="kpi-label">Tickets</div><div class="kpi-value" id="ka-tkt">—</div><div class="kpi-sub">Volumen de ventas</div></div>
    <div class="kpi" style="--ac:#4FA8D8"><div class="kpi-label">Ventas $ promedio por ticket</div><div class="kpi-value" id="ka-vtkt">—</div><div class="kpi-sub">Ventas $ ÷ Tickets</div></div>
    <div class="kpi" style="--ac:#7FD1EF"><div class="kpi-label">Unidades promedio por Ticket</div><div class="kpi-value" id="ka-utkt">—</div><div class="kpi-sub">Unidades ÷ Tickets</div></div>
  </div>
  <div class="tc">
    <div class="card-head">
      <div><div class="card-title">Ventas Históricas Mensuales</div><div class="card-sub">Acumulado por mes · Sucursales y meses seleccionados</div></div>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead><tr><th>Período</th><th class="r">Unidades</th><th class="r">Ventas $</th><th class="r">Utilidad</th><th class="r">Margen</th><th class="r">Tickets</th></tr></thead>
      <tbody id="tabla-historico"></tbody>
    </table>
    </div>
  </div>
  <div class="hist-card">
    <div class="card-head">
      <div><div class="card-title">Tendencia Histórica Mensual</div><div class="card-sub">Ventas mensuales y Tasa de incremento vs. mes anterior · Sucursales y meses seleccionados</div></div>
      <div class="metric-tabs">
        <button class="tab-btn active" id="btn-h-ventas"   onclick="changeHistMetric('ventas')">Ventas c/Desc</button>
        <button class="tab-btn"        id="btn-h-utilidad" onclick="changeHistMetric('utilidad')">Utilidad</button>
        <button class="tab-btn"        id="btn-h-unidades" onclick="changeHistMetric('unidades')">Unidades</button>
        <button class="tab-btn"        id="btn-h-tickets"  onclick="changeHistMetric('tickets')">Tickets</button>
      </div>
    </div>
    <div class="cw hist-chart"><canvas id="chart-historico"></canvas></div>
  </div>
  <hr class="section-divider">
  <div class="lines-container">
    <div class="lines-table-box">
      <div class="card-head">
        <div><div class="card-title">Resumen de Ventas Acumuladas por Línea - Categoría</div><div class="card-sub">Acumulado de ventas según selección</div></div>
      </div>
      <div class="table-scale-wrap">
        <table>
          <thead><tr><th style="width:26px"></th><th>Línea</th><th class="r">Unidades</th><th class="r">Ventas $</th><th class="r">Utilidad</th><th class="r">Margen</th></tr></thead>
          <tbody id="tabla-lineas"></tbody>
        </table>
      </div>
    </div>
    <div class="lines-chart-box">
      <div class="card-head" style="margin-bottom:.7rem">
        <div><div class="card-title">Tendencia de Categorías</div><div class="card-sub">Evolución mensual por categoría · Sucursales y meses seleccionados</div></div>
        <div class="metric-tabs">
          <button class="tab-btn active" id="btn-m-ventas"   onclick="changeLineMetric('ventas')">Ventas c/Desc</button>
          <button class="tab-btn"        id="btn-m-unidades" onclick="changeLineMetric('unidades')">Unidades</button>
          <button class="tab-btn"        id="btn-m-utilidad" onclick="changeLineMetric('utilidad')">Utilidad</button>
        </div>
      </div>
      <div class="cw" id="lineas-chart-wrap"><canvas id="chart-lineas"></canvas></div>
    </div>
  </div>
</div>
</div>
<div id="tab-toparticulos" class="tab-content">
<div class="main">
  <div class="top-kpi-row">
    <div class="top-kpi" style="--ac:#0B4F6C"><div class="kpi-label">Unidades vendidas Top 10</div><div class="kpi-value" id="tk-uni">—</div><div class="kpi-sub">Suma de top 10</div></div>
    <div class="top-kpi" style="--ac:#aa7300"><div class="kpi-label">Ventas $ Top 10</div><div class="kpi-value gold" id="tk-ventas">—</div><div class="kpi-sub">Suma de top 10</div></div>
    <div class="top-kpi" style="--ac:#1A7A4A"><div class="kpi-label">% participación sobre ventas $</div><div class="kpi-value" id="tk-margen" style="color:#1A7A4A">—</div><div class="kpi-sub">Ventas $ Top 10 ÷ Ventas $ acumuladas</div></div>
  </div>
  <div class="top-layout">
    <div class="top-table-box">
      <div class="card-head">
        <div><div class="card-title">Ranking de Artículos</div><div class="card-sub">Top 10 por Unidades Vendidas</div></div>
        <span class="note-bol">Acorde a selección</span>
      </div>
      <div class="table-scale-wrap">
      <table>
        <thead>
          <tr>
            <th style="width:36px;text-align:center">#</th>
            <th>Artículo</th>
            <th>Descripción</th>
            <th>Fabricante</th>
            <th class="r">Unidades</th>
            <th class="r">Venta $</th>
            <th class="r">% Venta</th>
          </tr>
        </thead>
        <tbody id="tabla-top"></tbody>
      </table>
      </div>
    </div>
    <div class="top-bubble-box">
      <div class="card-head" style="margin-bottom:.5rem">
        <div>
          <div class="card-title">Mapa de Desempeño · Top 10</div>
          <div class="card-sub">Eje X: Ventas $ · Eje Y: Margen (%) · Tamaño burbuja: Unidades vendidas</div>
        </div>
      </div>
      <div class="cw bubble-chart"><canvas id="chart-bubble"></canvas></div>
    </div>
  </div>
</div>
</div>
<div id="tab-lineascategoria" class="tab-content">
<div class="main">
  <div class="top-kpi-row">
    <div class="top-kpi" style="--ac:#00B0F0"><div class="kpi-label">Línea Líder</div><div class="kpi-value purple kv-txt" id="tk2-linea-nombre" style="font-size:1.05rem">—</div><div class="kpi-sub" id="tk2-linea-sub">—</div></div>
    <div class="top-kpi" style="--ac:#045C7F"><div class="kpi-label">Categoría Líder</div><div class="kpi-value kv-txt" id="tk2-cat-nombre" style="color:#045C7F;font-size:1.05rem">—</div><div class="kpi-sub" id="tk2-cat-sub">—</div></div>
    <div class="top-kpi" style="--ac:#1A7A4A"><div class="kpi-label">Margen · Categoría Líder</div><div class="kpi-value" id="tk2-cat-margen" style="color:#1A7A4A">—</div><div class="kpi-sub">Utilidad ÷ Ventas de la categoría líder</div></div>
  </div>
  <div class="top-table-box">
    <div class="card-head">
      <div><div class="card-title">Detalle por Línea y Categoría</div><div class="card-sub">Clic en una línea para desplegar sus categorías a detalle</div></div>
      <span class="note-bol">Acorde a sucursales y meses seleccionados</span>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead>
        <tr>
          <th style="width:26px"></th>
          <th>Línea / Categoría</th>
          <th class="r">Unidades</th>
          <th class="r">Ventas $</th>
          <th class="r">Utilidad</th>
          <th class="r">Margen</th>
        </tr>
      </thead>
      <tbody id="tabla-lineascategoria"></tbody>
    </table>
    </div>
  </div>
  <div class="hist-card" style="margin-top:1rem">
    <div class="card-head">
      <div><div class="card-title">Mapa de Línea &gt; Categoría</div><div class="card-sub">Top 3 categorías por línea </div></div>
      <span class="legend-gradient">Menor venta<span class="legend-bar"></span>Mayor venta</span>
    </div>
    <div class="cw" id="lc-chart-wrap" style="height:360px"><canvas id="chart-treemap"></canvas></div>
  </div>
</div>
</div>
<div id="tab-fabricantes" class="tab-content">
<div class="main">
  <div class="top-kpi-row">
    <div class="top-kpi" style="--ac:#00B0F0"><div class="kpi-label">Ventas $ · Fabricantes al 50%</div><div class="kpi-value purple" id="tk3-ventas">—</div><div class="kpi-sub">Suma de fabricantes hasta cubrir el 50%</div></div>
    <div class="top-kpi" style="--ac:#1A7A4A"><div class="kpi-label">Margen · Fabricantes al 50%</div><div class="kpi-value" id="tk3-margen" style="color:#1A7A4A">—</div><div class="kpi-sub">Utilidad ÷ Ventas del grupo</div></div>
    <div class="top-kpi" style="--ac:#aa7300"><div class="kpi-label">Fabricantes incluidos</div><div class="kpi-value gold" id="tk3-pct">—</div><div class="kpi-sub" id="tk3-pct-sub">Necesarios para alcanzar el 50% de ventas</div></div>
  </div>
  <div class="hist-card" style="margin-bottom:1rem">
    <div class="card-head">
      <div><div class="card-title">Fabricantes que concentran el 50% de las Ventas</div><div class="card-sub">% de participación en Ventas $</div></div>
    </div>
    <div class="cw" id="fab-chart-wrap" style="height:420px"><canvas id="chart-fabricantes"></canvas></div>
  </div>
  <div class="top-table-box">
    <div class="card-head">
      <div><div class="card-title">Detalle Completo por Fabricante</div><div class="card-sub">Todos los fabricantes de la selección, agrupados en cuartiles según su venta acumulada · clic en un cuartil para desplegar el detalle</div></div>
      <span class="note-bol">Acorde a sucursales y meses seleccionados</span>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead>
        <tr>
          <th style="width:26px"></th>
          <th>Fabricante</th>
          <th class="r">Ventas $</th>
          <th class="r">Margen</th>
          <th class="r">% Participación</th>
        </tr>
      </thead>
      <tbody id="tabla-fabricantes"></tbody>
    </table>
    </div>
  </div>
</div>
</div>
__TABS_CONTENT_EXTRA__
<footer style="text-align:center;padding:1.4rem 2rem 2rem;font-size:.7rem;color:#7391A0;border-top:1px solid #D9EEF7;margin-top:1rem">
  Elaborado por el Equipo de Planeación y Análisis de la información.
</footer>
<script>
document.addEventListener("DOMContentLoaded", function() {
    const RAW         = __DATA_JSON__;
    const LINEAS      = __LINEA_JSON__;
    const HISTORICO   = __HISTORICO_JSON__;
    const TOP_ART     = __TOP_ART_JSON__;
    const LINEAS_CAT  = __LINEAS_CAT_JSON__;
    const FABRICANTES = __FABRICANTES_JSON__;
    const PRESUPUESTO = __PRESUPUESTO_JSON__;
    const CATEGORIAS_DIARIO = __CATEGORIAS_DIARIO_JSON__;
    const SUCS        = __SUCURSALES_JSON__;
    const PERIODOS    = __PERIODOS_JSON__;
    const RESUMENES_ANT = __RESUMENES_ANT_JSON__;
    const CURRENT_PERIOD = __CURRENT_PERIOD_JSON__;
    // ── Segmentadores mutuamente excluyentes (estilo Excel) ──
    // Mapas de disponibilidad cruzada sucursal<->período, derivados de
    // HISTORICO. HISTORICO ya incluye las ventas del mes en curso (no sólo
    // meses cerrados), así que el cruce sucursal<->período es directo y no
    // necesita casos especiales para sucursales de apertura reciente: si una
    // sucursal sólo vendió en el mes en curso, aparecerá asociada únicamente
    // a ese período, y se deshabilitará igual que cualquier otro período
    // fuera de la selección activa.
    const SUCURSALES_POR_PERIODO = {};
    const PERIODOS_POR_SUCURSAL  = {};
    HISTORICO.forEach(r => {
        if(!(r.ventas > 0)) return;
        if(!SUCURSALES_POR_PERIODO[r.PeriodoLabel]) SUCURSALES_POR_PERIODO[r.PeriodoLabel] = new Set();
        SUCURSALES_POR_PERIODO[r.PeriodoLabel].add(r.NombreSucursal);
        if(!PERIODOS_POR_SUCURSAL[r.NombreSucursal]) PERIODOS_POR_SUCURSAL[r.NombreSucursal] = new Set();
        PERIODOS_POR_SUCURSAL[r.NombreSucursal].add(r.PeriodoLabel);
    });
    // Criterio de disponibilidad para 'Resumen Mes en Curso': esa pestaña no
    // usa HISTORICO/activeMeses (el segmentador de meses está oculto ahí),
    // así que el segmentador de sucursales debe regirse por RAW (venta real
    // del mes en curso), no por el histórico mensual.
    const RAW_SUCS = new Set(RAW.filter(r => r.ventas > 0).map(r => r.NombreSucursal));
    const MESES_ABR = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
    const PALETTE = [
        '#00B0F0','#14323D','#1E88B8','#4FA8D8','#045C7F',
        '#3AA0C9','#0B4F6C','#6FD3F7','#0D3B4F','#57B6DC',
        '#2B7A9E','#5BC0DE','#073042','#92DCF2','#0F6B8C',
        '#4695B8','#1D5266','#7FD1EF','#2F4858','#3D8FB0',
    ];
    const PALETTE_LINEAS = [
        '#045C7F','#00B0F0','#1E88B8','#4FA8D8','#6FD3F7',
        '#0D3B4F','#0B4F6C','#92DCF2','#0F3F52','#7FD1EF'
    ];
    const PALETTE_BUBBLE = [
        '#00B0F0','#14323D','#1E88B8','#045C7F','#4FA8D8',
        '#0B4F6C','#073042','#7FD1EF','#2B7A9E','#0D3B4F'
    ];
    function buildColorMap(list, palette) {
        const map = {};
        list.forEach((name, i) => { map[name] = palette[i % palette.length]; });
        return map;
    }
    const ALL_LINEAS = [...new Set(LINEAS.map(r => r.Línea))].sort();
    const LC = buildColorMap(ALL_LINEAS, PALETTE_LINEAS);
    // Colores por Categoría (para las gráficas de tendencia por categoría en
    // 'Resumen Mes Actual' y 'Resumen Histórico'). Se combinan las categorías
    // presentes en el histórico mensual y en el detalle diario del mes en
    // curso, más 'Otras categorías' (usado cuando hay más de TOP_N).
    const ALL_CATEGORIAS = [...new Set([
        ...LINEAS_CAT.map(r => r.Categoría),
        ...CATEGORIAS_DIARIO.map(r => r.Categoría),
        'Otras categorías'
    ])].sort();
    const CATC = buildColorMap(ALL_CATEGORIAS, PALETTE);
    let SC = {};
    const DAYS = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];
    let active      = new Set(SUCS);
    let activeMeses = new Set(PERIODOS);
    let charts      = {};
    let currentLineMetric = 'ventas';
    let currentLineMetricActual = 'ventas';
    let currentHistMetric = 'ventas';
    let currentSucMetric  = 'ventas';
    let currentTab = 'objetivos';
    const fM  = v => '$'+(v>=1e6?(v/1e6).toFixed(1)+'M':v>=1e3?(v/1e3).toFixed(0)+'K':Math.round(v));
    const fF  = v => '$'+Math.round(v).toLocaleString('es-MX');
    const fP  = v => (v*100).toFixed(2)+'%';
    const fN  = v => Math.round(v).toLocaleString('es-MX');
    const PREMIUM_TOOLTIP_OPTS = {
        backgroundColor:'#ffffff',titleColor:'#16232B',bodyColor:'#2B3A42',
        borderColor:'#D2ECF5',borderWidth:1,padding:10,cornerRadius:8,
        boxPadding:6,usePointStyle:true,
        titleFont:{family:"'Segoe UI', sans-serif",weight:'bold',size:12},
        bodyFont:{family:"'Segoe UI', sans-serif",size:12}
    };
    Chart.register(ChartDataLabels);
    function fitTables(){
        document.querySelectorAll('.table-scale-wrap').forEach(wrap => {
            const table = wrap.querySelector('table');
            if(!table) return;
            table.style.transform = 'none';
            wrap.style.height = 'auto';
            const availWidth  = wrap.clientWidth;
            const neededWidth = table.scrollWidth;
            if(availWidth > 0 && neededWidth > availWidth){
                const scale = availWidth / neededWidth;
                table.style.transform = `scale(${scale})`;
                wrap.style.height = (table.offsetHeight * scale) + 'px';
            }
        });
    }
    window.addEventListener('resize', () => {
        clearTimeout(window._fitTablesTimer);
        window._fitTablesTimer = setTimeout(fitTables, 150);
    });
    window.switchTab = function(tab) {
        currentTab = tab;
        document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
        document.querySelectorAll('.tab-nav-btn').forEach(el => el.classList.remove('active'));
        document.getElementById('tab-' + tab).classList.add('active');
        document.getElementById('tabnav-' + tab).classList.add('active');
        // El segmentador de "Meses" sólo aplica a pestañas que comparan varios
        // períodos; en 'objetivos' y 'resumenactual' (mes en curso) se oculta
        // para no sugerir que tiene efecto ahí (fuente de confusión previa).
        const usaMeses = (tab !== 'objetivos' && tab !== 'resumenactual');
        document.getElementById('filter-bar-mes').style.display = usaMeses ? 'flex' : 'none';
        // El segmentador de "Sucursales" tampoco afecta 'objetivos' (esa tabla
        // siempre muestra todas las sucursales, independientemente de la
        // selección), así que también se oculta ahí.
        const usaSucursales = (tab !== 'objetivos');
        document.getElementById('filter-bar-suc').style.display = usaSucursales ? 'flex' : 'none';
        // El criterio de "qué sucursal tiene datos" depende de la pestaña:
        // en 'resumenactual' se basa en el mes en curso (RAW); en el resto,
        // en los meses históricos seleccionados. Se recalcula al entrar.
        if(usaSucursales) refreshSucursalesDisponibilidad();
        if (tab === 'resumenactual') { update(); updateResumenActualExtra(); }
        if (tab === 'resumenacumulado') updateAcumulado();
        if (tab === 'objetivos') updateObjetivos();
        if (tab === 'toparticulos') updateTopArticulos();
        if (tab === 'lineascategoria') updateLineasCategoria();
        if (tab === 'fabricantes') updateFabricantes();
        if (tab.startsWith('resumen_')) updateResumenAnterior(tab.replace('resumen_',''));
        setTimeout(fitTables, 50);
    };
    function buildButtons(){
        SC = buildColorMap(SUCS, PALETTE);
        const wrap = document.getElementById('suc-btns');
        wrap.innerHTML = '';
        SUCS.forEach(s => {
            const btn = document.createElement('button');
            btn.className = 'suc-btn active';
            btn.dataset.sucursal = s;
            btn.innerHTML = `<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${SC[s]};margin-right:6px;flex-shrink:0"></span>${s}`;
            btn.addEventListener('click', () => {
                if(btn.disabled) return;
                if(active.has(s)){ active.delete(s); btn.classList.remove('active'); }
                else { active.add(s); btn.classList.add('active'); }
                refreshMesesDisponibilidad();
                updateAll();
            });
            wrap.appendChild(btn);
        });
    }
    function buildMesButtons(){
        const wrap = document.getElementById('mes-btns');
        wrap.innerHTML = '';
        PERIODOS.forEach(p => {
            const btn = document.createElement('button');
            btn.className = 'mes-btn active';
            btn.textContent = p;
            btn.dataset.periodo = p;
            btn.addEventListener('click', () => {
                if(btn.disabled) return;
                if(activeMeses.has(p)){ activeMeses.delete(p); btn.classList.remove('active'); }
                else { activeMeses.add(p); btn.classList.add('active'); }
                refreshSucursalesDisponibilidad();
                if(currentTab === 'resumenacumulado') updateAcumulado();
                if(currentTab === 'toparticulos') updateTopArticulos();
                if(currentTab === 'lineascategoria') updateLineasCategoria();
                if(currentTab === 'fabricantes') updateFabricantes();
            });
            wrap.appendChild(btn);
        });
    }
    function refreshSucursalesDisponibilidad(){
        let cambio = false;
        const criterioMesActual = (currentTab === 'resumenactual');
        document.querySelectorAll('#suc-btns .suc-btn').forEach(btn => {
            const s = btn.dataset.sucursal;
            let esValida;
            if(criterioMesActual){
                esValida = RAW_SUCS.has(s);
            } else {
                esValida = activeMeses.size === 0; // sin meses marcados = sin restricción
                if(!esValida){
                    for(const p of activeMeses){
                        if((SUCURSALES_POR_PERIODO[p] || new Set()).has(s)){ esValida = true; break; }
                    }
                }
            }
            btn.disabled = !esValida;
            btn.classList.toggle('oculto', !esValida);
            if(!esValida && active.has(s)){
                active.delete(s);
                btn.classList.remove('active');
                cambio = true;
            }
        });
        return cambio;
    }
    function refreshMesesDisponibilidad(){
        let cambio = false;
        document.querySelectorAll('#mes-btns .mes-btn').forEach(btn => {
            const p = btn.dataset.periodo;
            let esValido = active.size === 0; // sin sucursales marcadas = sin restricción
            if(!esValido){
                for(const s of active){
                    if((PERIODOS_POR_SUCURSAL[s] || new Set()).has(p)){ esValido = true; break; }
                }
            }
            btn.disabled = !esValido;
            btn.classList.toggle('oculto', !esValido);
            if(!esValido && activeMeses.has(p)){
                activeMeses.delete(p);
                btn.classList.remove('active');
                cambio = true;
            }
        });
        return cambio;
    }
    window.toggleAll = function(){
        const btns = [...document.querySelectorAll('#suc-btns .suc-btn')];
        const seleccionables = btns.filter(b => !b.disabled);
        const todasActivas = seleccionables.length>0 && seleccionables.every(b => b.classList.contains('active'));
        if(todasActivas){
            seleccionables.forEach(b => { active.delete(b.dataset.sucursal); b.classList.remove('active'); });
        } else {
            seleccionables.forEach(b => { active.add(b.dataset.sucursal); b.classList.add('active'); });
        }
        refreshMesesDisponibilidad();
        updateAll();
    };
    window.toggleAllMeses = function(){
        const btns = [...document.querySelectorAll('#mes-btns .mes-btn')];
        const seleccionables = btns.filter(b => !b.disabled);
        const todosActivos = seleccionables.length>0 && seleccionables.every(b => b.classList.contains('active'));
        if(todosActivos){
            seleccionables.forEach(b => { activeMeses.delete(b.dataset.periodo); b.classList.remove('active'); });
        } else {
            seleccionables.forEach(b => { activeMeses.add(b.dataset.periodo); b.classList.add('active'); });
        }
        refreshSucursalesDisponibilidad();
        if(currentTab === 'resumenacumulado') updateAcumulado();
        if(currentTab === 'toparticulos') updateTopArticulos();
        if(currentTab === 'lineascategoria') updateLineasCategoria();
        if(currentTab === 'fabricantes') updateFabricantes();
    };
    function updateAll(){
        if(currentTab === 'resumenactual') { update(); updateResumenActualExtra(); }
        if(currentTab === 'resumenacumulado') updateAcumulado();
        if(currentTab === 'toparticulos') updateTopArticulos();
        if(currentTab === 'lineascategoria') updateLineasCategoria();
        if(currentTab === 'fabricantes') updateFabricantes();
        if(currentTab && currentTab.startsWith('resumen_')) updateResumenAnterior(currentTab.replace('resumen_',''));
    }
    function getFiltered()          { return RAW.filter(r => active.has(r.NombreSucursal)); }
    function getFilteredLineas()    {
        return LINEAS.filter(r =>
            active.has(r.NombreSucursal) && activeMeses.has(r.PeriodoLabel)
        );
    }
    function getFilteredLineasActual() {
        return LINEAS.filter(r =>
            active.has(r.NombreSucursal) && r.PeriodoLabel === CURRENT_PERIOD
        );
    }
    function getFilteredHist()      {
        return HISTORICO.filter(r =>
            active.has(r.NombreSucursal) && activeMeses.has(r.PeriodoLabel)
        );
    }
    function getFilteredTopArt()    {
        return TOP_ART.filter(r =>
            active.has(r.NombreSucursal) && activeMeses.has(r.PeriodoLabel)
        );
    }
    function getFilteredLineasCat() {
        return LINEAS_CAT.filter(r =>
            active.has(r.NombreSucursal) && activeMeses.has(r.PeriodoLabel)
        );
    }
    function getFilteredFabricantes() {
        return FABRICANTES.filter(r =>
            active.has(r.NombreSucursal) && activeMeses.has(r.PeriodoLabel)
        );
    }
    function getFilteredCategoriasDiario() {
        // Sólo filtra por sucursal: este dataset ya viene acotado al mes en
        // curso real (no aplica el segmentador de meses, oculto en esta pestaña).
        return CATEGORIAS_DIARIO.filter(r => active.has(r.NombreSucursal));
    }
    function salesColor(t){
        const clamped = Math.max(0, Math.min(1, t));
        const c1 = [207,236,248], c2 = [6,60,84];
        const r = Math.round(c1[0] + (c2[0]-c1[0])*clamped);
        const g = Math.round(c1[1] + (c2[1]-c1[1])*clamped);
        const b = Math.round(c1[2] + (c2[2]-c1[2])*clamped);
        return `rgba(${r},${g},${b},0.92)`;
    }
    // Degradado morado (misma paleta del layout) en vez de colores categóricos
    // por sucursal: escala mejor cuando se agregan más sucursales y mantiene
    // identidad visual con el resto del dashboard. t=0 -> lavanda claro,
    // t=1 -> morado oscuro (#00B0F0-ish).
    function purpleGradientColor(t, alpha){
        const clamped = Math.max(0, Math.min(1, t));
        const c1 = [207,236,248], c2 = [6,60,84];
        const r = Math.round(c1[0] + (c2[0]-c1[0])*clamped);
        const g = Math.round(c1[1] + (c2[1]-c1[1])*clamped);
        const b = Math.round(c1[2] + (c2[2]-c1[2])*clamped);
        return `rgba(${r},${g},${b},${alpha})`;
    }
    function purpleGradientColors(values, alpha){
        const max = Math.max(...values, 0);
        const min = Math.min(...values, 0);
        return values.map(v => {
            const t = max>min ? (v-min)/(max-min) : 1;
            return purpleGradientColor(0.15 + t*0.85, alpha);
        });
    }
    function getDates()      { return [...new Set(RAW.map(r => r.FechaStr))].sort(); }
    function getActiveSucs() { return [...active].sort(); }
    function aggByDate(data){
        const m = {};
        data.forEach(r => {
            if(!m[r.FechaStr]) m[r.FechaStr] = {unidades:0,ventas:0,utilidad:0,tickets:0};
            m[r.FechaStr].unidades  += r.unidades;
            m[r.FechaStr].ventas    += r.ventas;
            m[r.FechaStr].utilidad  += r.utilidad;
            m[r.FechaStr].tickets   += r.tickets;
        });
        return m;
    }
    function aggBySucursal(data){
        const m = {};
        data.forEach(r => {
            if(!m[r.NombreSucursal]) m[r.NombreSucursal] = {NombreSucursal:r.NombreSucursal, unidades:0, ventas:0, utilidad:0, tickets:0};
            m[r.NombreSucursal].unidades += r.unidades;
            m[r.NombreSucursal].ventas   += r.ventas;
            m[r.NombreSucursal].utilidad += r.utilidad;
            m[r.NombreSucursal].tickets  += r.tickets;
        });
        return Object.values(m);
    }
    function dc(id){ if(charts[id]){ charts[id].destroy(); delete charts[id]; } }
    window.changeLineMetric = function(metric){
        currentLineMetric = metric;
        document.querySelectorAll('[id^="btn-m-"]').forEach(b => b.classList.remove('active'));
        document.getElementById('btn-m-' + metric).classList.add('active');
        updateLineas();
    };
    window.changeLineMetricActual = function(metric){
        currentLineMetricActual = metric;
        document.querySelectorAll('[id^="btn-ma-"]').forEach(b => b.classList.remove('active'));
        document.getElementById('btn-ma-' + metric).classList.add('active');
        updateLineasActual();
    };
    window.changeHistMetric = function(metric){
        currentHistMetric = metric;
        document.querySelectorAll('[id^="btn-h-"]').forEach(b => b.classList.remove('active'));
        document.getElementById('btn-h-' + metric).classList.add('active');
        updateHistorico();
    };
    window.changeSucMetric = function(metric){
        currentSucMetric = metric;
        document.querySelectorAll('[id^="btn-hs-"]').forEach(b => b.classList.remove('active'));
        document.getElementById('btn-hs-' + metric).classList.add('active');
        updateChartSucursalActual();
    };
    function objetivosData(){
        return PRESUPUESTO.slice().sort((a,b) => a.ClaveSucursal - b.ClaveSucursal);
    }
    function updateObjetivos(){
        updatePresupuesto();
        updateChartObjetivos();
        fitTables();
    }
    function updatePresupuesto(){
        const data  = objetivosData();
        const tbody = document.getElementById('tabla-presupuesto');
        const kpiRow = document.getElementById('obj-kpi-row');
        if(!data.length){
            tbody.innerHTML = '<tr><td colspan="8" class="empty">No hay sucursales con presupuesto ni venta en el mes en curso.</td></tr>';
            ['p-ventas','p-pronostico','p-presupuesto','p-cumplimiento'].forEach(id => { document.getElementById(id).textContent = '—'; });
            document.querySelectorAll('#th-presupuesto, #th-alcance, [data-col="presupuesto"], [data-col="alcance"]').forEach(el => { el.style.display = 'none'; });
            ['kpi-presupuesto','kpi-cumplimiento'].forEach(id => { const el = document.getElementById(id); if(el) el.style.display = 'none'; });
            if(kpiRow) kpiRow.style.gridTemplateColumns = 'repeat(2,1fr)';
            return;
        }
        let tU=0, tV=0, tUt=0, tP=0, tB=0, tVpres=0;
        const rows = data.map(r => {
            tU += r.unidadesActual; tV += r.ventasActual; tUt += r.utilidadActual;
            tP += r.pronostico;
            // Presupuesto/cumplimiento se dejan EN BLANCO (no en cero) cuando la
            // sucursal no tiene presupuesto asignado para el mes en curso.
            const tienePresupuesto = r.presupuesto !== null && r.presupuesto !== undefined && r.presupuesto > 0;
            if(tienePresupuesto){ tB += r.presupuesto; tVpres += r.ventasActual; }
            const cump  = tienePresupuesto ? r.ventasActual/r.presupuesto : null;
            const cls   = tienePresupuesto ? (cump>=0.98 ? 'hi' : cump>=0.90 ? 'mi' : 'lo') : '';
            const mgCls = r.margen>=0.40 ? 'hi' : 'mi';
            return `<tr>
                <td data-label="Sucursal"><b>${r.NombreSucursal}</b></td>
                <td class="r" data-label="Unidades">${fN(r.unidadesActual)}</td>
                <td class="r" data-label="Venta $">${fF(r.ventasActual)}</td>
                <td class="r" data-label="Utilidad">${fF(r.utilidadActual)}</td>
                <td class="r" data-label="Margen"><span class="pill ${mgCls}">${fP(r.margen)}</span></td>
                <td class="r" data-label="Presupuesto" data-col="presupuesto">${tienePresupuesto ? fF(r.presupuesto) : ''}</td>
                <td class="r" data-label="Alcance actual" data-col="alcance">${tienePresupuesto ? `<span class="pill ${cls}">${fP(cump)}</span>` : '<span style="color:#A9C6D3">—</span>'}</td>
                <td class="r" data-label="Pronóstico de Alcance"><b>${fF(r.pronostico)}</b></td>
            </tr>`;
        }).join('');
        // El total de Presupuesto y Cumplimiento se calculan SOLO con las
        // sucursales que sí tienen presupuesto asignado; las que no lo tienen
        // no se convierten en cero ni distorsionan este acumulado.
        const totMg   = tV>0 ? tUt/tV : 0;
        const totCump = tB>0 ? tVpres/tB : 0;
        const totCls  = totCump>=0.98 ? 'hi' : totCump>=0.90 ? 'mi' : 'lo';
        const totPresTexto      = tB>0 ? fF(tB) : '';
        const totCumplimientoTd = tB>0 ? `<span class="pill ${totCls}">${fP(totCump)}</span>` : '<span style="color:#A9C6D3">—</span>';
        tbody.innerHTML = rows + `<tr class="total-row">
            <td data-label=""><b>TOTAL GENERAL</b></td>
            <td class="r" data-label="Unidades">${fN(tU)}</td>
            <td class="r" data-label="Venta $">${fF(tV)}</td>
            <td class="r" data-label="Utilidad">${fF(tUt)}</td>
            <td class="r" data-label="Margen"><span class="pill ${totMg>=.40?'hi':'mi'}">${fP(totMg)}</span></td>
            <td class="r" data-label="Presupuesto" data-col="presupuesto">${totPresTexto}</td>
            <td class="r" data-label="Cumplimiento" data-col="alcance">${totCumplimientoTd}</td>
            <td class="r" data-label="Pronóstico de Alcance"><b>${fF(tP)}</b></td>
        </tr>`;
        document.getElementById('p-ventas').textContent       = fF(tV);
        document.getElementById('p-pronostico').textContent   = fF(tP);
        document.getElementById('p-presupuesto').textContent  = tB>0 ? fF(tB) : '—';
        document.getElementById('p-cumplimiento').textContent = tB>0 ? fP(totCump) : '—';
        // ── Ocultar columnas/tarjetas de Presupuesto y Alcance actual si
        // ninguna sucursal tiene presupuesto asignado (aún no hay intención
        // de asignarlo en el corto plazo). ──
        const hasPresupuesto = tB > 0;
        document.querySelectorAll('#th-presupuesto, #th-alcance, [data-col="presupuesto"], [data-col="alcance"]').forEach(el => {
            el.style.display = hasPresupuesto ? '' : 'none';
        });
        const elPres = document.getElementById('kpi-presupuesto');
        const elCump = document.getElementById('kpi-cumplimiento');
        if(elPres) elPres.style.display = hasPresupuesto ? '' : 'none';
        if(elCump) elCump.style.display = hasPresupuesto ? '' : 'none';
        if(kpiRow) kpiRow.style.gridTemplateColumns = hasPresupuesto ? 'repeat(4,1fr)' : 'repeat(2,1fr)';
    }
    function updateChartObjetivos(){
        dc('objetivos');
        const data = objetivosData();
        if(!data.length) return;
        document.getElementById('obj-chart-wrap').style.height =
            Math.max(360, Math.min(620, 340 + data.length * 6)) + 'px';
        const hasPresupuesto = data.some(r => r.presupuesto !== null && r.presupuesto !== undefined && r.presupuesto > 0);
        const maxVal = Math.max(
            ...data.map(r => Math.max(r.ventasActual, r.presupuesto || 0, r.pronostico || 0))
        ) || 0;
        const datasets = [
            {
                type: 'bar',
                label: 'Ventas $ mes en curso',
                data: data.map(r => r.ventasActual),
                backgroundColor: '#00B0F0cc',
                borderColor: '#00B0F0',
                borderWidth: 1,
                borderRadius: 4,
            }
        ];
        if(hasPresupuesto){
            datasets.push({
                type: 'bar',
                label: 'Presupuesto',
                data: data.map(r => (r.presupuesto !== null && r.presupuesto !== undefined && r.presupuesto > 0) ? r.presupuesto : null),
                backgroundColor: '#aa7300aa',
                borderColor: '#aa7300',
                borderWidth: 1,
                borderRadius: 4,
            });
        }
        datasets.push({
            type: 'line',
            label: 'Pronóstico de Alcance',
            data: data.map(r => r.pronostico),
            borderColor: '#1A7A4A',
            backgroundColor: 'transparent',
            borderWidth: 2,
            borderDash: [6, 4],
            pointRadius: 3,
            pointHoverRadius: 5,
            pointBackgroundColor: '#1A7A4A',
            tension: 0.1,
            fill: false,
            datalabels: { display: false },
        });
        charts['objetivos'] = new Chart(document.getElementById('chart-objetivos'), {
            data: {
                labels: data.map(r => r.NombreSucursal),
                datasets
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    legend: {
                        display: true, position: 'top', onClick: null,
                        labels: {
                            boxWidth: 9, boxHeight: 9, usePointStyle: true, pointStyle: 'circle',
                            padding: 12, color: '#46626D',
                            font: { family: "'Segoe UI', sans-serif", size: 10, weight: '600' }
                        }
                    },
                    datalabels: { display: false },
                    tooltip: {
                        ...PREMIUM_TOOLTIP_OPTS,
                        callbacks: {
                            label: ctx => ctx.raw === null ? null : ` ${ctx.dataset.label}: ${fF(ctx.raw)}`,
                            footer: items => {
                                if(!items.length) return '';
                                const r = data[items[0].dataIndex];
                                if(!(r.presupuesto > 0)) return 'Sin presupuesto asignado';
                                const avance = r.ventasActual / r.presupuesto;
                                return `Avance real: ${fP(avance)}`;
                            }
                        },
                        footerColor: '#00B0F0',
                        footerFont: { family: "'Segoe UI', sans-serif", size: 11, weight: 'bold' }
                    }
                },
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: { font: { size: 9 }, color: '#666666', maxRotation: 60, minRotation: 45, autoSkip: false }
                    },
                    y: {
                        grid: { color: '#EAF7FC' },
                        suggestedMax: maxVal * 1.1,
                        ticks: { font: { size: 9 }, color: '#666666', callback: v => fM(v) }
                    }
                }
            }
        });
    }
    function update(){
        const data       = getFiltered();
        const dates      = getDates();
        const byDate     = aggByDate(data);
        const activeSucs = getActiveSucs();
        const multi      = activeSucs.length > 1;
        const totV  = data.reduce((a,r) => a+r.ventas,   0);
        const totU  = data.reduce((a,r) => a+r.utilidad, 0);
        const totUn = data.reduce((a,r) => a+r.unidades, 0);
        const totTk = data.reduce((a,r) => a+r.tickets,  0);
        document.getElementById('k-uni').textContent      = fN(totUn);
        document.getElementById('k-ventas').textContent   = fF(totV);
        document.getElementById('k-ventas-s').textContent = activeSucs.length + ' sucursal(es)';
        document.getElementById('k-util').textContent     = fF(totU);
        document.getElementById('k-util-sub').textContent = 'Margen: ' + fP(totV>0 ? totU/totV : 0);
        document.getElementById('k-tkt').textContent      = fN(totTk);
        document.getElementById('k-vtkt').textContent     = totTk>0 ? fF(totV/totTk) : '—';
        document.getElementById('k-utkt').textContent     = totTk>0 ? (totUn/totTk).toFixed(1) : '—';
        updateTabla(dates, byDate);
        updateChart('ventas',  dates, byDate, activeSucs, multi);
        updateChart('tickets', dates, byDate, activeSucs, multi);
        fitTables();
    }
    function updateTabla(dates, byDate){
        const tbody = document.getElementById('tabla-body');
        const valid = dates.filter(d => byDate[d]);
        if(!valid.length){
            tbody.innerHTML = '<tr><td colspan="7" class="empty">Selecciona sucursales para mapear la grilla.</td></tr>';
            return;
        }
        let tUn=0,tV=0,tU=0,tTk=0;
        const rows = valid.map(d => {
            const r  = byDate[d];
            const mg = r.ventas>0 ? r.utilidad/r.ventas : 0;
            const pill = mg>=.40 ? 'hi' : 'mi';
            const dn = DAYS[new Date(d+'T12:00:00').getDay()];
            const mAbr = MESES_ABR[parseInt(d.slice(5,7),10)-1];
            tUn+=r.unidades; tV+=r.ventas; tU+=r.utilidad; tTk+=r.tickets;
            return `<tr><td class="date" data-label="Fecha">${d.slice(8)} ${mAbr}</td><td class="dayname" data-label="Día">${dn}</td><td class="r" data-label="Unidades">${fN(r.unidades)}</td><td class="r" data-label="Ventas $"><b>${fF(r.ventas)}</b></td><td class="r" data-label="Utilidad">${fF(r.utilidad)}</td><td class="r" data-label="Margen"><span class="pill ${pill}">${fP(mg)}</span></td><td class="r" data-label="Tickets"><b>${fN(r.tickets)}</b></td></tr>`;
        }).join('');
        const totMg = tV>0 ? tU/tV : 0;
        tbody.innerHTML = rows + `<tr class="total-row"><td data-label="" colspan="2"><b>TOTAL PERÍODO</b></td><td class="r" data-label="Unidades">${fN(tUn)}</td><td class="r" data-label="Ventas $">${fF(tV)}</td><td class="r" data-label="Utilidad">${fF(tU)}</td><td class="r" data-label="Margen"><span class="pill ${totMg>=.40?'hi':'mi'}">${fP(totMg)}</span></td><td class="r" data-label="Tickets">${fN(tTk)}</td></tr>`;
    }
    function updateChartGeneric(chartId, field, sourceData, dates, byDate, activeSucs, multi){
        dc(chartId);
        const canvasEl = document.getElementById('chart-'+chartId);
        if(!canvasEl) return;
        const labels    = dates.map(d => d.slice(8));
        const isVentas  = field === 'ventas';
        const monoColor = isVentas ? '#00B0F0' : '#14323D';
        const datasets  = multi
            ? activeSucs.map(s => ({
                label: s,
                data: dates.map(d => { const r = sourceData.find(x => x.FechaStr===d && x.NombreSucursal===s); return r ? r[field] : 0; }),
                borderColor: SC[s]||'#888', borderWidth:1.6,
                pointRadius:3.5, pointHitRadius:20, pointHoverRadius:6, tension:0.12, fill:false
              }))
            : [{
                label: isVentas ? 'Ventas c/Desc' : 'Cantidad de Tickets',
                data: dates.map(d => byDate[d] ? byDate[d][field] : 0),
                borderColor: monoColor, backgroundColor: monoColor+'08', borderWidth:2.2,
                pointRadius:3.5, pointHitRadius:20, pointHoverRadius:6, tension:0.12, fill:true
              }];
        charts[chartId] = new Chart(canvasEl, {
            type:'line', data:{labels, datasets},
            options:{
                responsive:true, maintainAspectRatio:false,
                interaction:{mode:'index', intersect:false},
                plugins:{
                    legend:{display:multi, position:'top', onClick:null,
                        labels:{boxWidth:8,boxHeight:8,usePointStyle:true,pointStyle:'circle',padding:12,color:'#46626D',font:{family:"'Segoe UI', sans-serif",size:10,weight:'600'}}},
                    datalabels:{display:false},
                    tooltip:PREMIUM_TOOLTIP_OPTS
                },
                scales:{
                    x:{grid:{display:false}, ticks:{font:{size:10},color:'#666666'}},
                    y:{grid:{color:'#EAF7FC'}, ticks:{font:{size:10},color:'#666666', callback: isVentas ? v=>fM(v) : v=>v}}
                }
            }
        });
    }
    function updateChart(field, dates, byDate, activeSucs, multi){
        updateChartGeneric(field, field, RAW, dates, byDate, activeSucs, multi);
    }
    function updateLineas(){
        dc('lineas');
        const tbody    = document.getElementById('tabla-lineas');
        const filtered = getFilteredLineasCat();
        if(!filtered.length){
            tbody.innerHTML = '<tr><td colspan="6" class="empty">Selecciona sucursales y meses para desplegar líneas.</td></tr>';
            document.getElementById('lineas-chart-wrap').style.height = '220px';
            fitTables();
            return;
        }
        // ── Tabla expandible: Línea (total) -> Categorías (detalle) ──
        const catMap  = {};   // Línea|||Categoría -> totales
        const lineMap = {};   // Línea -> totales + lista de categorías
        let totalUnidades=0, totalVentas=0, totalUtilidad=0;
        filtered.forEach(r => {
            const k = r.Línea + '|||' + r.Categoría;
            if(!catMap[k]) catMap[k] = {Línea:r.Línea, Categoría:r.Categoría, unidades:0, ventas:0, utilidad:0};
            catMap[k].unidades += r.unidades;
            catMap[k].ventas   += r.ventas;
            catMap[k].utilidad += r.utilidad;
            if(!lineMap[r.Línea]) lineMap[r.Línea] = {Línea:r.Línea, unidades:0, ventas:0, utilidad:0, cats:[]};
            lineMap[r.Línea].unidades += r.unidades;
            lineMap[r.Línea].ventas   += r.ventas;
            lineMap[r.Línea].utilidad += r.utilidad;
            totalUnidades += r.unidades; totalVentas += r.ventas; totalUtilidad += r.utilidad;
        });
        Object.values(catMap).forEach(c => { lineMap[c.Línea].cats.push(c); });
        const lineasSorted = Object.values(lineMap).sort((a,b) => b[currentLineMetric]-a[currentLineMetric]);
        let html = '';
        lineasSorted.forEach((l, i) => {
            const mg  = l.ventas>0 ? l.utilidad/l.ventas : 0;
            const rid = 'lcH'+i;
            html += `<tr class="lc-parent" onclick="toggleLineaCat('${rid}')">
                <td data-label="" style="text-align:right"><span class="lc-icon" id="${rid}-icon">▸</span></td>
                <td data-label="Línea"><span style="display:inline-block;width:8px;height:8px;border-radius:4px;background:${LC[l.Línea]||'#888'};margin-right:7px"></span><b>${l.Línea}</b></td>
                <td class="r" data-label="Unidades">${fN(l.unidades)}</td>
                <td class="r" data-label="Ventas $"><b>${fF(l.ventas)}</b></td>
                <td class="r" data-label="Utilidad">${fF(l.utilidad)}</td>
                <td class="r" data-label="Margen"><span class="pill ${mg>=.40?'hi':'mi'}">${fP(mg)}</span></td>
            </tr>`;
            const catsSorted = l.cats.slice().sort((a,b) => b.ventas-a.ventas);
            catsSorted.forEach(c => {
                const cmg = c.ventas>0 ? c.utilidad/c.ventas : 0;
                html += `<tr class="lc-child ${rid}-child" style="display:none">
                    <td data-label=""></td>
                    <td data-label="Categoría">${c.Categoría}</td>
                    <td class="r" data-label="Unidades">${fN(c.unidades)}</td>
                    <td class="r" data-label="Ventas $">${fF(c.ventas)}</td>
                    <td class="r" data-label="Utilidad">${fF(c.utilidad)}</td>
                    <td class="r" data-label="Margen"><span class="pill ${cmg>=.40?'hi':'mi'}">${fP(cmg)}</span></td>
                </tr>`;
            });
        });
        const totalMargen = totalVentas>0 ? totalUtilidad/totalVentas : 0;
        tbody.innerHTML = html + `<tr class="total-row"><td data-label=""></td><td data-label=""><b>TOTAL ACUMULADO</b></td><td class="r" data-label="Unidades">${fN(totalUnidades)}</td><td class="r" data-label="Ventas $">${fF(totalVentas)}</td><td class="r" data-label="Utilidad">${fF(totalUtilidad)}</td><td class="r" data-label="Margen"><span class="pill ${totalMargen>=.40?'hi':'mi'}">${fP(totalMargen)}</span></td></tr>`;
        // ── Gráfico de línea: tendencia mensual por categoría ──
        const periodosOrdenados = PERIODOS.filter(p => activeMeses.has(p));
        const catTotals = {};
        Object.values(catMap).forEach(c => { catTotals[c.Categoría] = (catTotals[c.Categoría]||0) + c[currentLineMetric]; });
        const TOP_N = 8;
        const catEntries = Object.entries(catTotals).sort((a,b) => b[1]-a[1]);
        const topCats  = catEntries.slice(0, TOP_N).map(e => e[0]);
        const topSet   = new Set(topCats);
        const hasOtras = catEntries.length > TOP_N;
        const seriesNames = hasOtras ? [...topCats, 'Otras categorías'] : topCats;
        const byPeriodCat = {};
        seriesNames.forEach(n => byPeriodCat[n] = {});
        filtered.forEach(r => {
            const catName = topSet.has(r.Categoría) ? r.Categoría : (hasOtras ? 'Otras categorías' : r.Categoría);
            if(!byPeriodCat[catName]) byPeriodCat[catName] = {};
            byPeriodCat[catName][r.PeriodoLabel] = (byPeriodCat[catName][r.PeriodoLabel]||0) + r[currentLineMetric];
        });
        const isMoneda = currentLineMetric==='ventas' || currentLineMetric==='utilidad';
        const datasets = seriesNames.map(name => {
            const color = CATC[name] || '#888';
            return {
                label: name,
                data: periodosOrdenados.map(p => byPeriodCat[name][p] || 0),
                borderColor: color, backgroundColor: color+'15',
                borderWidth: 2, pointRadius: 3, pointHitRadius: 16, pointHoverRadius: 5,
                tension: 0.15, fill: false,
            };
        });
        document.getElementById('lineas-chart-wrap').style.height = '320px';
        charts['lineas'] = new Chart(document.getElementById('chart-lineas'), {
            type:'line',
            data:{ labels: periodosOrdenados, datasets },
            options:{
                responsive:true, maintainAspectRatio:false,
                interaction:{mode:'index', intersect:false},
                plugins:{
                    legend:{display:true, position:'top', onClick:null,
                        labels:{boxWidth:8,boxHeight:8,usePointStyle:true,pointStyle:'circle',padding:10,color:'#46626D',font:{family:"'Segoe UI', sans-serif",size:9.5,weight:'600'}}},
                    datalabels:{display:false},
                    tooltip:{...PREMIUM_TOOLTIP_OPTS, callbacks:{label:ctx=>` ${ctx.dataset.label}: ${isMoneda?fF(ctx.raw):fN(ctx.raw)}`}}
                },
                scales:{
                    x:{grid:{display:false}, ticks:{font:{size:9},color:'#666666'}},
                    y:{grid:{color:'#EAF7FC'}, ticks:{font:{size:9},color:'#666666', callback: isMoneda ? v=>fM(v) : v=>fN(v)}}
                }
            }
        });
        fitTables();
    }
    function updateLineasActual(){
        dc('lineasactual');
        const tbody    = document.getElementById('tabla-lineas-actual');
        const filtered = getFilteredCategoriasDiario();
        if(!filtered.length){
            tbody.innerHTML = '<tr><td colspan="6" class="empty">Selecciona sucursales para desplegar líneas.</td></tr>';
            document.getElementById('lineasactual-chart-wrap').style.height = '220px';
            return;
        }
        // ── Tabla expandible: Línea (total del mes) -> Categorías (detalle) ──
        const catMap  = {};   // Línea|||Categoría -> totales (sumando todos los días)
        const lineMap = {};   // Línea -> totales + lista de categorías
        let totalUnidades=0, totalVentas=0, totalUtilidad=0;
        filtered.forEach(r => {
            const k = r.Línea + '|||' + r.Categoría;
            if(!catMap[k]) catMap[k] = {Línea:r.Línea, Categoría:r.Categoría, unidades:0, ventas:0, utilidad:0};
            catMap[k].unidades += r.unidades;
            catMap[k].ventas   += r.ventas;
            catMap[k].utilidad += r.utilidad;
            if(!lineMap[r.Línea]) lineMap[r.Línea] = {Línea:r.Línea, unidades:0, ventas:0, utilidad:0, cats:[]};
            lineMap[r.Línea].unidades += r.unidades;
            lineMap[r.Línea].ventas   += r.ventas;
            lineMap[r.Línea].utilidad += r.utilidad;
            totalUnidades += r.unidades; totalVentas += r.ventas; totalUtilidad += r.utilidad;
        });
        Object.values(catMap).forEach(c => { lineMap[c.Línea].cats.push(c); });
        const lineasSorted = Object.values(lineMap).sort((a,b) => b[currentLineMetricActual]-a[currentLineMetricActual]);
        let html = '';
        lineasSorted.forEach((l, i) => {
            const mg  = l.ventas>0 ? l.utilidad/l.ventas : 0;
            const rid = 'lcA'+i;
            html += `<tr class="lc-parent" onclick="toggleLineaCat('${rid}')">
                <td data-label="" style="text-align:right"><span class="lc-icon" id="${rid}-icon">▸</span></td>
                <td data-label="Línea"><span style="display:inline-block;width:8px;height:8px;border-radius:4px;background:${LC[l.Línea]||'#888'};margin-right:7px"></span><b>${l.Línea}</b></td>
                <td class="r" data-label="Unidades">${fN(l.unidades)}</td>
                <td class="r" data-label="Ventas $"><b>${fF(l.ventas)}</b></td>
                <td class="r" data-label="Utilidad">${fF(l.utilidad)}</td>
                <td class="r" data-label="Margen"><span class="pill ${mg>=.40?'hi':'mi'}">${fP(mg)}</span></td>
            </tr>`;
            const catsSorted = l.cats.slice().sort((a,b) => b.ventas-a.ventas);
            catsSorted.forEach(c => {
                const cmg = c.ventas>0 ? c.utilidad/c.ventas : 0;
                html += `<tr class="lc-child ${rid}-child" style="display:none">
                    <td data-label=""></td>
                    <td data-label="Categoría">${c.Categoría}</td>
                    <td class="r" data-label="Unidades">${fN(c.unidades)}</td>
                    <td class="r" data-label="Ventas $">${fF(c.ventas)}</td>
                    <td class="r" data-label="Utilidad">${fF(c.utilidad)}</td>
                    <td class="r" data-label="Margen"><span class="pill ${cmg>=.40?'hi':'mi'}">${fP(cmg)}</span></td>
                </tr>`;
            });
        });
        const totalMargen = totalVentas>0 ? totalUtilidad/totalVentas : 0;
        tbody.innerHTML = html + `<tr class="total-row"><td data-label=""></td><td data-label=""><b>TOTAL MES EN CURSO</b></td><td class="r" data-label="Unidades">${fN(totalUnidades)}</td><td class="r" data-label="Ventas $">${fF(totalVentas)}</td><td class="r" data-label="Utilidad">${fF(totalUtilidad)}</td><td class="r" data-label="Margen"><span class="pill ${totalMargen>=.40?'hi':'mi'}">${fP(totalMargen)}</span></td></tr>`;
        // ── Gráfico de línea: tendencia diaria por categoría ──
        const dates = getDates();
        const catTotals = {};
        Object.values(catMap).forEach(c => { catTotals[c.Categoría] = (catTotals[c.Categoría]||0) + c[currentLineMetricActual]; });
        const TOP_N = 8;
        const catEntries = Object.entries(catTotals).sort((a,b) => b[1]-a[1]);
        const topCats  = catEntries.slice(0, TOP_N).map(e => e[0]);
        const topSet   = new Set(topCats);
        const hasOtras = catEntries.length > TOP_N;
        const seriesNames = hasOtras ? [...topCats, 'Otras categorías'] : topCats;
        const byDateCat = {};
        seriesNames.forEach(n => byDateCat[n] = {});
        filtered.forEach(r => {
            const catName = topSet.has(r.Categoría) ? r.Categoría : (hasOtras ? 'Otras categorías' : r.Categoría);
            if(!byDateCat[catName]) byDateCat[catName] = {};
            byDateCat[catName][r.FechaStr] = (byDateCat[catName][r.FechaStr]||0) + r[currentLineMetricActual];
        });
        const isMoneda = currentLineMetricActual==='ventas' || currentLineMetricActual==='utilidad';
        const datasets = seriesNames.map(name => {
            const color = CATC[name] || '#888';
            return {
                label: name,
                data: dates.map(d => byDateCat[name][d] || 0),
                borderColor: color, backgroundColor: color+'15',
                borderWidth: 2, pointRadius: 2.5, pointHitRadius: 14, pointHoverRadius: 5,
                tension: 0.15, fill: false,
            };
        });
        document.getElementById('lineasactual-chart-wrap').style.height = '320px';
        charts['lineasactual'] = new Chart(document.getElementById('chart-lineasactual'), {
            type:'line',
            data:{ labels: dates.map(d => d.slice(8)), datasets },
            options:{
                responsive:true, maintainAspectRatio:false,
                interaction:{mode:'index', intersect:false},
                plugins:{
                    legend:{display:true, position:'top', onClick:null,
                        labels:{boxWidth:8,boxHeight:8,usePointStyle:true,pointStyle:'circle',padding:10,color:'#46626D',font:{family:"'Segoe UI', sans-serif",size:9.5,weight:'600'}}},
                    datalabels:{display:false},
                    tooltip:{...PREMIUM_TOOLTIP_OPTS, callbacks:{label:ctx=>` ${ctx.dataset.label}: ${isMoneda?fF(ctx.raw):fN(ctx.raw)}`}}
                },
                scales:{
                    x:{grid:{display:false}, ticks:{font:{size:9},color:'#666666'}},
                    y:{grid:{color:'#EAF7FC'}, ticks:{font:{size:9},color:'#666666', callback: isMoneda ? v=>fM(v) : v=>fN(v)}}
                }
            }
        });
    }
    function updateTablaSucursalActual(){
        const tbody = document.getElementById('tabla-sucursal-actual');
        const data  = aggBySucursal(getFiltered()).sort((a,b) => b.ventas - a.ventas);
        if(!data.length){
            tbody.innerHTML = '<tr><td colspan="6" class="empty">Selecciona sucursales para ver el comparativo.</td></tr>';
            return;
        }
        let tUn=0,tV=0,tU=0,tTk=0;
        const rows = data.map(r => {
            const mg = r.ventas>0 ? r.utilidad/r.ventas : 0;
            tUn+=r.unidades; tV+=r.ventas; tU+=r.utilidad; tTk+=r.tickets;
            return `<tr><td data-label="Sucursal"><span style="display:inline-block;width:8px;height:8px;border-radius:4px;background:${SC[r.NombreSucursal]||'#888'};margin-right:7px"></span><b>${r.NombreSucursal}</b></td><td class="r" data-label="Unidades">${fN(r.unidades)}</td><td class="r" data-label="Ventas $"><b>${fF(r.ventas)}</b></td><td class="r" data-label="Utilidad">${fF(r.utilidad)}</td><td class="r" data-label="Margen"><span class="pill ${mg>=.40?'hi':'mi'}">${fP(mg)}</span></td><td class="r" data-label="Tickets">${fN(r.tickets)}</td></tr>`;
        }).join('');
        const totMg = tV>0 ? tU/tV : 0;
        tbody.innerHTML = rows + `<tr class="total-row"><td data-label=""><b>TOTAL MES EN CURSO</b></td><td class="r" data-label="Unidades">${fN(tUn)}</td><td class="r" data-label="Ventas $">${fF(tV)}</td><td class="r" data-label="Utilidad">${fF(tU)}</td><td class="r" data-label="Margen"><span class="pill ${totMg>=.40?'hi':'mi'}">${fP(totMg)}</span></td><td class="r" data-label="Tickets">${fN(tTk)}</td></tr>`;
    }
    function updateChartSucursalActual(){
        dc('sucursalactual');
        const data = aggBySucursal(getFiltered()).sort((a,b) => b[currentSucMetric]-a[currentSucMetric]);
        const wrap = document.getElementById('sucursalactual-chart-wrap');
        if(!data.length){ if(wrap) wrap.style.height = '220px'; return; }
        const isMoneda  = currentSucMetric==='ventas' || currentSucMetric==='utilidad';
        const margenes  = data.map(r => r.ventas>0 ? +(r.utilidad/r.ventas*100).toFixed(2) : 0);
        const metricVals = data.map(r => r[currentSucMetric]);
        const barFill    = purpleGradientColors(metricVals, 0.85);
        const barBorder  = purpleGradientColors(metricVals, 1);
        if(wrap) wrap.style.height = Math.max(260, data.length*40 + 90) + 'px';
        charts['sucursalactual'] = new Chart(document.getElementById('chart-sucursalactual'), {
            data: {
                labels: data.map(r => r.NombreSucursal),
                datasets: [
                    {
                        type:'bar',
                        label: currentSucMetric==='ventas'?'Ventas c/Desc':currentSucMetric==='utilidad'?'Utilidad':currentSucMetric==='unidades'?'Unidades':'Tickets',
                        data: data.map(r => r[currentSucMetric]),
                        backgroundColor: barFill,
                        borderColor: barBorder,
                        borderWidth:1, borderRadius:5, yAxisID:'y'
                    },
                    {
                        type:'line', label:'Margen %', yAxisID:'y2',
                        data: margenes, borderColor:'#aa7300', backgroundColor:'transparent',
                        borderWidth:2, pointRadius:4, pointHoverRadius:6, tension:0.15,
                        datalabels:{display:false}
                    }
                ]
            },
            options:{
                responsive:true, maintainAspectRatio:false,
                interaction:{mode:'index', intersect:false},
                plugins:{
                    legend:{display:true, position:'top', onClick:null,
                        labels:{boxWidth:8,boxHeight:8,usePointStyle:true,pointStyle:'circle',padding:10,color:'#46626D',font:{family:"'Segoe UI', sans-serif",size:10,weight:'600'}}},
                    datalabels:{display:false},
                    tooltip:{...PREMIUM_TOOLTIP_OPTS, callbacks:{
                        label: ctx => ctx.dataset.yAxisID==='y2' ? ` Margen: ${ctx.raw.toFixed(2)}%` : ` ${ctx.dataset.label}: ${isMoneda?fF(ctx.raw):fN(ctx.raw)}`
                    }}
                },
                scales:{
                    x:{grid:{display:false}, ticks:{font:{size:9},color:'#666666',maxRotation:45}},
                    y:{position:'left', grid:{color:'#EAF7FC'}, ticks:{font:{size:9},color:'#666666', callback: v => isMoneda?fM(v):fN(v)}},
                    y2:{position:'right', grid:{display:false}, ticks:{font:{size:9},color:'#aa7300', callback: v => v.toFixed(1)+'%'}}
                }
            }
        });
    }
    function updateResumenActualExtra(){
        updateTablaSucursalActual();
        updateChartSucursalActual();
        updateLineasActual();
        fitTables();
    }
    function updateKPIsAcumulado(){
        const data  = getFilteredHist();
        const totU  = data.reduce((a,r) => a+r.unidades, 0);
        const totV  = data.reduce((a,r) => a+r.ventas,   0);
        const totUt = data.reduce((a,r) => a+r.utilidad, 0);
        const totTk = data.reduce((a,r) => a+r.tickets,  0);
        document.getElementById('ka-uni').textContent      = fN(totU);
        document.getElementById('ka-ventas').textContent   = fF(totV);
        document.getElementById('ka-ventas-s').textContent = getActiveSucs().length + ' sucursal(es) · ' + activeMeses.size + ' período(s)';
        document.getElementById('ka-util').textContent     = fF(totUt);
        document.getElementById('ka-util-sub').textContent = 'Margen: ' + fP(totV>0 ? totUt/totV : 0);
        document.getElementById('ka-tkt').textContent      = fN(totTk);
        document.getElementById('ka-vtkt').textContent     = totTk>0 ? fF(totV/totTk) : '—';
        document.getElementById('ka-utkt').textContent     = totTk>0 ? (totU/totTk).toFixed(1) : '—';
    }
    function updateAcumulado(){
        updateKPIsAcumulado();
        updateHistorico();
        updateLineas();
        fitTables();
    }
    function updateHistorico(){
        dc('historico');
        const tbody    = document.getElementById('tabla-historico');
        const filtered = getFilteredHist();
        if(!filtered.length){
            tbody.innerHTML = '<tr><td colspan="6" class="empty">Selecciona sucursales y meses para ver el histórico.</td></tr>';
            fitTables();
            return;
        }
        const periodoMap   = {};
        const periodoOrder = [];
        filtered.forEach(r => {
            if(!periodoMap[r.PeriodoLabel]){
                periodoMap[r.PeriodoLabel] = {PeriodoLabel:r.PeriodoLabel, Año:r.Año, MesNum:r.MesNum,
                    unidades:0, ventas:0, utilidad:0, tickets:0};
                periodoOrder.push(r.PeriodoLabel);
            }
            periodoMap[r.PeriodoLabel].unidades  += r.unidades;
            periodoMap[r.PeriodoLabel].ventas    += r.ventas;
            periodoMap[r.PeriodoLabel].utilidad  += r.utilidad;
            periodoMap[r.PeriodoLabel].tickets   += r.tickets;
        });
        const sorted = periodoOrder
            .map(p => periodoMap[p])
            .sort((a,b) => a.Año!==b.Año ? a.Año-b.Año : a.MesNum-b.MesNum);
        let tUn=0,tV=0,tU=0,tTk=0;
        sorted.forEach(r => { tUn+=r.unidades; tV+=r.ventas; tU+=r.utilidad; tTk+=r.tickets; });
        const totMg = tV>0 ? tU/tV : 0;
        const rows = sorted.map(r => {
            const mg = r.ventas>0 ? r.utilidad/r.ventas : 0;
            return `<tr><td class="date" data-label="Período">${r.PeriodoLabel}</td><td class="r" data-label="Unidades">${fN(r.unidades)}</td><td class="r" data-label="Ventas $"><b>${fF(r.ventas)}</b></td><td class="r" data-label="Utilidad">${fF(r.utilidad)}</td><td class="r" data-label="Margen"><span class="pill ${mg>=.40?'hi':'mi'}">${fP(mg)}</span></td><td class="r" data-label="Tickets"><b>${fN(r.tickets)}</b></td></tr>`;
        }).join('');
        tbody.innerHTML = rows + `<tr class="total-row"><td data-label=""><b>TOTAL HISTÓRICO</b></td><td class="r" data-label="Unidades">${fN(tUn)}</td><td class="r" data-label="Ventas $">${fF(tV)}</td><td class="r" data-label="Utilidad">${fF(tU)}</td><td class="r" data-label="Margen"><span class="pill ${totMg>=.40?'hi':'mi'}">${fP(totMg)}</span></td><td class="r" data-label="Tickets">${fN(tTk)}</td></tr>`;
        const activeSucs = getActiveSucs();
        const multi      = activeSucs.length > 1;
        const isMoneda   = currentHistMetric==='ventas' || currentHistMetric==='utilidad';
        const labels     = sorted.map(r => r.PeriodoLabel);
        // Tasa de incremento mensual: SIEMPRE contra el mes calendario
        // inmediato anterior real (tomado de HISTORICO sin filtrar por
        // 'activeMeses'), no contra el período anterior visible en la
        // gráfica. Así, si sólo hay Nov y Jun seleccionados, Nov se compara
        // contra Oct real y Jun contra Mayo real, aunque no estén marcados.
        const crecimientos = sorted.map(r => {
            let prevMes = r.MesNum - 1, prevAnio = r.Año;
            if(prevMes === 0){ prevMes = 12; prevAnio -= 1; }
            const prevTotal = HISTORICO
                .filter(x => active.has(x.NombreSucursal) && x.Año===prevAnio && x.MesNum===prevMes)
                .reduce((a,x) => a + x[currentHistMetric], 0);
            if(prevTotal <= 0) return null; // sin dato real del mes anterior -> no se puede calcular
            return +(((r[currentHistMetric] - prevTotal) / prevTotal) * 100).toFixed(2);
        });
        const barDatasets = multi
            ? activeSucs.map(s => {
                const color = SC[s]||'#888';
                return {
                    type:'bar', label:s, yAxisID:'y',
                    data: sorted.map(p => {
                        return filtered
                            .filter(r => r.NombreSucursal===s && r.PeriodoLabel===p.PeriodoLabel)
                            .reduce((a,r) => a+r[currentHistMetric], 0);
                    }),
                    backgroundColor: color+'cc', borderColor: color,
                    borderWidth: 1, borderRadius: 3, stack: 'stack0',
                };
              })
            : [{
                type:'bar',
                label: currentHistMetric==='ventas'?'Ventas c/Desc':currentHistMetric==='utilidad'?'Utilidad':currentHistMetric==='unidades'?'Unidades':'Tickets',
                data: sorted.map(r => r[currentHistMetric]),
                backgroundColor:'#00B0F0cc', borderColor:'#00B0F0',
                borderWidth:1, borderRadius:4, yAxisID:'y',
              }];
        const crecimientoLine = {
            type:'line', label:'Crecimiento % (vs. mes anterior)', yAxisID:'y2',
            data: crecimientos,
            borderColor:'#aa7300', backgroundColor:'transparent',
            borderWidth:2, pointRadius:3.5, pointHoverRadius:6, tension:0.2,
            spanGaps:false,
            datalabels:{display:false}
        };
        charts['historico'] = new Chart(document.getElementById('chart-historico'), {
            data:{ labels, datasets:[...barDatasets, crecimientoLine] },
            options:{
                responsive:true, maintainAspectRatio:false,
                interaction:{mode:'index', intersect:false},
                plugins:{
                    legend:{display:true, position:'top', onClick:null,
                        labels:{boxWidth:8,boxHeight:8,usePointStyle:true,pointStyle:'circle',padding:10,color:'#46626D',
                            font:{family:"'Segoe UI', sans-serif",size:10,weight:'600'}}},
                    datalabels:{display:false},
                    tooltip:{...PREMIUM_TOOLTIP_OPTS, callbacks:{
                        label: ctx => {
                            if(ctx.dataset.yAxisID==='y2'){
                                if(ctx.raw === null || ctx.raw === undefined) return ' Crecimiento: sin dato del mes anterior';
                                const signo = ctx.raw >= 0 ? '+' : '';
                                return ` Crecimiento vs. mes anterior: ${signo}${ctx.raw.toFixed(2)}%`;
                            }
                            return isMoneda ? ` ${ctx.dataset.label}: ${fF(ctx.raw)}` : ` ${ctx.dataset.label}: ${fN(ctx.raw)}`;
                        }
                    }}
                },
                scales:{
                    x:{grid:{display:false}, ticks:{font:{size:9},color:'#666666',maxRotation:45}},
                    y:{stacked:true, position:'left', grid:{color:'#EAF7FC'}, ticks:{font:{size:9},color:'#666666', callback: v => isMoneda?fM(v):fN(v)}},
                    y2:{position:'right', grid:{display:false}, ticks:{font:{size:9},color:'#aa7300', callback: v => v.toFixed(1)+'%'}}
                }
            }
        });
        fitTables();
    }
    function updateTopArticulos(){
        dc('bubble');
        const filtered = getFilteredTopArt();
        if(!filtered.length){
            document.getElementById('tabla-top').innerHTML =
                '<tr><td colspan="7" class="empty">Selecciona sucursales y meses para ver el ranking.</td></tr>';
            ['tk-uni','tk-ventas','tk-margen'].forEach(id => {
                document.getElementById(id).textContent = '—';
            });
            fitTables();
            return;
        }
        const artMap = {};
        filtered.forEach(r => {
            if(!artMap[r.Artículo]) artMap[r.Artículo] = {
                Artículo:    r.Artículo,
                Descripción: r.Descripción,
                Fabricante:  r.Fabricante,
                unidades: 0,
                ventas:   0,
                utilidad: 0,
            };
            artMap[r.Artículo].unidades += r.unidades;
            artMap[r.Artículo].ventas   += r.ventas;
            artMap[r.Artículo].utilidad += r.utilidad;
        });
        const top10 = Object.values(artMap)
            .sort((a,b) => b.unidades - a.unidades)
            .slice(0, 10);
        const histFiltrado    = getFilteredHist();
        const totalHistVentas = histFiltrado.reduce((a,r) => a + r.ventas, 0);
        const totUni    = top10.reduce((a,r) => a+r.unidades, 0);
        const totVentas = top10.reduce((a,r) => a+r.ventas,   0);
        document.getElementById('tk-uni').textContent    = fN(totUni);
        document.getElementById('tk-ventas').textContent = fF(totVentas);
        document.getElementById('tk-margen').textContent = fP(totalHistVentas>0 ? totVentas/totalHistVentas : 0);
        const tbody = document.getElementById('tabla-top');
        const rows = top10.map((r, i) => {
            const rank = i + 1;
            const rc   = rank===1?'r1':rank===2?'r2':rank===3?'r3':'';
            const pct  = totalHistVentas>0 ? r.ventas/totalHistVentas : 0;
            return `<tr>
              <td data-label="#" style="text-align:right"><span class="badge-rank ${rc}">${rank}</span></td>
              <td class="art-code" data-label="Artículo">${r.Artículo}</td>
              <td class="art-desc" data-label="Descripción">${r.Descripción}</td>
              <td class="art-fab" data-label="Fabricante">${r.Fabricante}</td>
              <td class="r" data-label="Unidades"><b>${fN(r.unidades)}</b></td>
              <td class="r" data-label="Venta $">${fF(r.ventas)}</td>
              <td class="r" data-label="% Venta"><span class="pill hi">${fP(pct)}</span></td>
            </tr>`;
        }).join('');
        tbody.innerHTML = rows;
        const maxUni = Math.max(...top10.map(r => r.unidades)) || 1;
        const minUni = Math.min(...top10.map(r => r.unidades)) || 0;
        const minR = 10, maxR = 44;
        const rScale = v => {
            const t       = (v - minUni) / (maxUni - minUni || 1);
            const areaMin = minR * minR, areaMax = maxR * maxR;
            return Math.sqrt(areaMin + t * (areaMax - areaMin));
        };
        const bubbleDatasets = top10.map((r, i) => {
            const margen = r.ventas > 0 ? +(r.utilidad / r.ventas * 100).toFixed(2) : 0;
            return {
                label: r.Artículo,
                data: [{ x: r.ventas, y: margen, r: rScale(r.unidades) }],
                backgroundColor: PALETTE_BUBBLE[i % PALETTE_BUBBLE.length] + 'bb',
                borderColor:     PALETTE_BUBBLE[i % PALETTE_BUBBLE.length],
                borderWidth: 1.5,
                _meta: { desc: r.Descripción, fab: r.Fabricante, uni: r.unidades },
            };
        });
        charts['bubble'] = new Chart(document.getElementById('chart-bubble'), {
            type: 'bubble',
            data: { datasets: bubbleDatasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'nearest', intersect: true },
                plugins: {
                    legend: {
                        display: true,
                        position: 'bottom',
                        onClick: null,
                        labels: {
                            boxWidth: 9, boxHeight: 9, usePointStyle: true, pointStyle: 'circle',
                            padding: 12, color: '#46626D',
                            font: { family: "'Segoe UI', sans-serif", size: 9.5, weight: '600' }
                        }
                    },
                    datalabels: { display: false },
                    tooltip: {
                        ...PREMIUM_TOOLTIP_OPTS,
                        callbacks: {
                            title: ctx => ctx[0].dataset.label,
                            label: ctx => {
                                const ds  = ctx.dataset;
                                const pt  = ctx.raw;
                                return [
                                    ` ${ds._meta.desc}`,
                                    ` Fabricante: ${ds._meta.fab}`,
                                    ` Unidades: ${fN(ds._meta.uni)}`,
                                    ` Ventas: ${fF(pt.x)}`,
                                    ` Margen: ${pt.y.toFixed(2)}%`,
                                ];
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        title: { display: true, text: 'Ventas $ c/Desc', color: '#666666', font: { size: 10, weight: '600' } },
                        grid: { color: '#EAF7FC' },
                        ticks: { font: { size: 9 }, color: '#666666', callback: v => fM(v) }
                    },
                    y: {
                        title: { display: true, text: 'Margen %', color: '#666666', font: { size: 10, weight: '600' } },
                        grid: { color: '#EAF7FC' },
                        ticks: { font: { size: 9 }, color: '#666666', callback: v => v.toFixed(1)+'%' }
                    }
                }
            }
        });
        fitTables();
    }
    window.toggleLineaCat = function(rid){
        const rows = document.querySelectorAll('.'+rid+'-child');
        if(!rows.length) return;
        const show = rows[0].style.display === 'none';
        rows.forEach(r => { r.style.display = show ? 'table-row' : 'none'; });
        const icon = document.getElementById(rid+'-icon');
        if(icon) icon.textContent = show ? '▾' : '▸';
        fitTables();
    };
    window.toggleFabQuartil = function(rid){
        const rows = document.querySelectorAll('.'+rid+'-child');
        if(!rows.length) return;
        const show = rows[0].style.display === 'none';
        rows.forEach(r => { r.style.display = show ? 'table-row' : 'none'; });
        const icon = document.getElementById(rid+'-icon');
        if(icon) icon.textContent = show ? '▾' : '▸';
        fitTables();
    };
    function updateLineasCategoria(){
        dc('treemap');
        const tbody    = document.getElementById('tabla-lineascategoria');
        const filtered = getFilteredLineasCat();
        if(!filtered.length){
            tbody.innerHTML = '<tr><td colspan="6" class="empty">Selecciona sucursales y meses para ver el análisis.</td></tr>';
            document.getElementById('tk2-linea-nombre').textContent = '—';
            document.getElementById('tk2-linea-sub').textContent    = '—';
            document.getElementById('tk2-cat-nombre').textContent   = '—';
            document.getElementById('tk2-cat-sub').textContent      = '—';
            document.getElementById('tk2-cat-margen').textContent   = '—';
            fitTables();
            return;
        }
        const key2 = {};
        const lineaTotales = {};
        filtered.forEach(r => {
            const k = r.Línea + '|||' + r.Categoría;
            if(!key2[k]) key2[k] = {Línea:r.Línea, Categoría:r.Categoría, ventas:0, utilidad:0, unidades:0};
            key2[k].ventas    += r.ventas;
            key2[k].utilidad  += r.utilidad;
            key2[k].unidades  += r.unidades;
            lineaTotales[r.Línea] = (lineaTotales[r.Línea]||0) + r.ventas;
        });
        const allLeaves = Object.values(key2);
        const totalSeleccion = filtered.reduce((a,r) => a+r.ventas, 0);
        const lineaLiderEntry = Object.entries(lineaTotales).sort((a,b) => b[1]-a[1])[0];
        if(lineaLiderEntry){
            const [lineaLiderNombre, lineaLiderVentas] = lineaLiderEntry;
            const lineaLiderPct = totalSeleccion>0 ? lineaLiderVentas/totalSeleccion : 0;
            document.getElementById('tk2-linea-nombre').textContent = lineaLiderNombre;
            document.getElementById('tk2-linea-sub').textContent    = `${fF(lineaLiderVentas)} · ${fP(lineaLiderPct)} de participación`;
        }
        const catLiderEntry = allLeaves.slice().sort((a,b) => b.ventas - a.ventas)[0];
        if(catLiderEntry){
            const catLiderPct = totalSeleccion>0 ? catLiderEntry.ventas/totalSeleccion : 0;
            const catLiderMg  = catLiderEntry.ventas>0 ? catLiderEntry.utilidad/catLiderEntry.ventas : 0;
            document.getElementById('tk2-cat-nombre').textContent = catLiderEntry.Categoría;
            document.getElementById('tk2-cat-sub').textContent    = `${catLiderEntry.Línea} · ${fF(catLiderEntry.ventas)} · ${fP(catLiderPct)} de participación`;
            document.getElementById('tk2-cat-margen').textContent = fP(catLiderMg);
        }
        const lineaMap = {};
        allLeaves.forEach(r => {
            if(!lineaMap[r.Línea]) lineaMap[r.Línea] = {Línea:r.Línea, ventas:0, utilidad:0, unidades:0, cats:[]};
            lineaMap[r.Línea].ventas   += r.ventas;
            lineaMap[r.Línea].utilidad += r.utilidad;
            lineaMap[r.Línea].unidades += r.unidades;
            lineaMap[r.Línea].cats.push(r);
        });
        const lineasSorted = Object.values(lineaMap).sort((a,b) => b.ventas-a.ventas);
        let html = '';
        lineasSorted.forEach((l, i) => {
            const mg  = l.ventas>0 ? l.utilidad/l.ventas : 0;
            const rid = 'lc'+i;
            html += `<tr class="lc-parent" onclick="toggleLineaCat('${rid}')">
                <td data-label="" style="text-align:right"><span class="lc-icon" id="${rid}-icon">▸</span></td>
                <td data-label="Línea"><b>${l.Línea}</b></td>
                <td class="r" data-label="Unidades">${fN(l.unidades)}</td>
                <td class="r" data-label="Ventas $"><b>${fF(l.ventas)}</b></td>
                <td class="r" data-label="Utilidad">${fF(l.utilidad)}</td>
                <td class="r" data-label="Margen"><span class="pill ${mg>=.40?'hi':'mi'}">${fP(mg)}</span></td>
            </tr>`;
            const catsSorted = l.cats.slice().sort((a,b) => b.ventas-a.ventas);
            catsSorted.forEach(c => {
                const cmg = c.ventas>0 ? c.utilidad/c.ventas : 0;
                html += `<tr class="lc-child ${rid}-child" style="display:none">
                    <td data-label=""></td>
                    <td data-label="Categoría">${c.Categoría}</td>
                    <td class="r" data-label="Unidades">${fN(c.unidades)}</td>
                    <td class="r" data-label="Ventas $">${fF(c.ventas)}</td>
                    <td class="r" data-label="Utilidad">${fF(c.utilidad)}</td>
                    <td class="r" data-label="Margen"><span class="pill ${cmg>=.40?'hi':'mi'}">${fP(cmg)}</span></td>
                </tr>`;
            });
        });
        tbody.innerHTML = html;
        const treemapLeaves = [];
        Object.keys(lineaTotales).forEach(lineaName => {
            const topCats = allLeaves
                .filter(r => r.Línea === lineaName)
                .sort((a,b) => b.ventas - a.ventas)
                .slice(0, 3);
            treemapLeaves.push(...topCats);
        });
        const treemapLineaTotales = {};
        treemapLeaves.forEach(r => { treemapLineaTotales[r.Línea] = (treemapLineaTotales[r.Línea]||0) + r.ventas; });
        const lineaVals = Object.values(treemapLineaTotales);
        const leafVals  = treemapLeaves.map(r => r.ventas);
        const lineaMin  = Math.min(...lineaVals), lineaMax = Math.max(...lineaVals);
        const leafMin   = Math.min(...leafVals),  leafMax  = Math.max(...leafVals);
        const headerValueMap = {};
        Object.keys(treemapLineaTotales).forEach(lineaName => {
            headerValueMap[Math.round(treemapLineaTotales[lineaName])] = lineaName;
        });
        const numCeldas = treemapLeaves.length;
        const alturaTreemap = Math.max(360, Math.min(900, 360 + Math.max(0, numCeldas - 40) * 3));
        document.getElementById('lc-chart-wrap').style.height = alturaTreemap + 'px';
        charts['treemap'] = new Chart(document.getElementById('chart-treemap'), {
            type: 'treemap',
            data: {
                datasets: [{
                    label: 'Ventas c/Desc',
                    tree: treemapLeaves,
                    key: 'ventas',
                    groups: ['Línea','Categoría'],
                    spacing: 1,
                    borderWidth: 1.5,
                    borderColor: '#ffffff',
                    displayMode: 'headerBoxes',
                    captions: {
                        display: true,
                        color: '#ffffff',
                        font: { weight: '700', size: 10 },
                        formatter: ctx => ctx.raw.g
                    },
                    labels: {
                        display: ctx => ctx.type === 'data' && ctx.raw.l > 0 && ctx.raw.w > 46 && ctx.raw.h > 18,
                        color: ctx => {
                            if(ctx.type !== 'data' || ctx.raw.l === 0) return '#ffffff';
                            const v = ctx.raw.v || 0;
                            const t = leafMax > leafMin ? (v-leafMin)/(leafMax-leafMin) : 1;
                            return t < 0.45 ? '#073042' : '#ffffff';
                        },
                        font: { size: 8.5, weight: '600' },
                        overflow: 'hidden',
                        formatter: ctx => {
                            if(ctx.type !== 'data' || ctx.raw.l === 0) return '';
                            return ctx.raw.g;
                        }
                    },
                    backgroundColor: ctx => {
                        if(ctx.type !== 'data') return 'transparent';
                        const v = ctx.raw.v || 0;
                        if(ctx.raw.l === 0){
                            const t = lineaMax > lineaMin ? (v-lineaMin)/(lineaMax-lineaMin) : 1;
                            return salesColor(0.55 + t*0.45);
                        }
                        const t = leafMax > leafMin ? (v-leafMin)/(leafMax-leafMin) : 1;
                        return salesColor(t);
                    },
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    datalabels: { display: false},
                    legend: { display: false },
                    tooltip: {
                        ...PREMIUM_TOOLTIP_OPTS,
                        callbacks: {
                            title: ctx => {
                                const raw = ctx[0] && ctx[0].raw;
                                if(!raw) return '';
                                if(raw.l === 0){
                                    const nombre = raw.g || headerValueMap[Math.round(raw.v || raw.s || 0)] || '';
                                    return nombre ? `Línea: ${nombre}` : 'Línea';
                                }
                                const orig = raw._data;
                                if(orig && orig.Línea && orig.Categoría) return `${orig.Línea} · ${orig.Categoría}`;
                                return raw.g || 'Categoría';
                            },
                            label: ctx => ' Ventas: ' + fF(ctx.raw.v || 0)
                        }
                    }
                }
            }
        });
        fitTables();
    }
    const QUART_LABELS = {1:'Cuartil 1 · Mayor venta', 2:'Cuartil 2', 3:'Cuartil 3', 4:'Cuartil 4 · Menor venta'};
    const QUART_COLORS = {
        1: {bg:'#00B0F0', fg:'#ffffff', tint:'#00B0F014'},
        2: {bg:'#1E88B8', fg:'#ffffff', tint:'#1E88B812'},
        3: {bg:'#AEE3F7', fg:'#073042', tint:'#AEE3F720'},
        4: {bg:'#EAF7FC', fg:'#46626D', tint:'#EAF7FC40'},
    };
    function updateFabricantes(){
        dc('fabricantes');
        const tbody    = document.getElementById('tabla-fabricantes');
        const filtered = getFilteredFabricantes();
        if(!filtered.length){
            tbody.innerHTML = '<tr><td colspan="5" class="empty">Selecciona sucursales y meses para ver los fabricantes.</td></tr>';
            ['tk3-ventas','tk3-margen','tk3-pct'].forEach(id => { document.getElementById(id).textContent = '—'; });
            document.getElementById('tk3-pct-sub').textContent = 'Necesarios para alcanzar el 50% de ventas';
            fitTables();
            return;
        }
        const fabMap = {};
        let totalSeleccion = 0;
        filtered.forEach(r => {
            if(!fabMap[r.Fabricante]) fabMap[r.Fabricante] = {Fabricante:r.Fabricante, ventas:0, utilidad:0, unidades:0};
            fabMap[r.Fabricante].ventas   += r.ventas;
            fabMap[r.Fabricante].utilidad += r.utilidad;
            fabMap[r.Fabricante].unidades += r.unidades;
            totalSeleccion += r.ventas;
        });
        const sorted = Object.values(fabMap).sort((a,b) => b.ventas - a.ventas);
        let cum = 0;
        sorted.forEach(f => {
            const start = totalSeleccion>0 ? cum/totalSeleccion : 0;
            cum += f.ventas;
            f.pct    = totalSeleccion>0 ? f.ventas/totalSeleccion : 0;
            f.cumPct = totalSeleccion>0 ? cum/totalSeleccion : 0;
            if(start < 0.25)      f.cuartil = 1;
            else if(start < 0.50) f.cuartil = 2;
            else if(start < 0.75) f.cuartil = 3;
            else                  f.cuartil = 4;
        });
        const paretoList = [];
        let cumPareto = 0;
        for(const f of sorted){
            paretoList.push(f);
            cumPareto += f.ventas;
            if(totalSeleccion>0 && cumPareto/totalSeleccion >= 0.5) break;
        }
        const totVentasPareto = paretoList.reduce((a,r) => a+r.ventas,   0);
        const totUtilPareto   = paretoList.reduce((a,r) => a+r.utilidad, 0);
        document.getElementById('tk3-ventas').textContent = fF(totVentasPareto);
        document.getElementById('tk3-margen').textContent = fP(totVentasPareto>0 ? totUtilPareto/totVentasPareto : 0);
        document.getElementById('tk3-pct').textContent    = `${paretoList.length} de ${sorted.length}`;
        document.getElementById('tk3-pct-sub').textContent = `Cubren el ${fP(totalSeleccion>0 ? totVentasPareto/totalSeleccion : 0)} de las ventas`;
        const groups = {1:[], 2:[], 3:[], 4:[]};
        sorted.forEach(f => groups[f.cuartil].push(f));
        const maxVentas = sorted.length ? sorted[0].ventas : 0;
        let html = '';
        let cumBeforeGroup = 0;
        [1,2,3,4].forEach(q => {
            const items = groups[q];
            if(!items.length) return;
            const qVentas = items.reduce((a,r) => a+r.ventas,   0);
            const qUtil   = items.reduce((a,r) => a+r.utilidad, 0);
            const qMg     = qVentas>0 ? qUtil/qVentas : 0;
            const qPct    = totalSeleccion>0 ? qVentas/totalSeleccion : 0;
            const rid     = 'fq'+q;
            const qc      = QUART_COLORS[q];
            const expanded = q === 1;
            const startPct = totalSeleccion>0 ? cumBeforeGroup/totalSeleccion : 0;
            cumBeforeGroup += qVentas;
            const endPct   = totalSeleccion>0 ? cumBeforeGroup/totalSeleccion : 0;
            html += `<tr class="fq-parent" onclick="toggleFabQuartil('${rid}')" style="background:${qc.tint}">
                <td data-label="" style="text-align:right"><span class="lc-icon" id="${rid}-icon">${expanded?'▾':'▸'}</span></td>
                <td data-label="Cuartil"><span class="fq-badge" style="background:${qc.bg};color:${qc.fg}">${QUART_LABELS[q]}</span>
                    <span style="color:#666666;font-size:.68rem;margin-left:6px;display:block;text-align:left;margin-top:4px">${items.length} fabricante${items.length!==1?'s':''} · ${fP(startPct)}–${fP(endPct)} acumulado</span></td>
                <td class="r" data-label="Ventas $"><b>${fF(qVentas)}</b></td>
                <td class="r" data-label="Margen"><span class="pill ${qMg>=.40?'hi':'mi'}">${fP(qMg)}</span></td>
                <td class="r" data-label="% Participación"><b>${fP(qPct)}</b></td>
            </tr>`;
            items.forEach((f, idx) => {
                const mg  = f.ventas>0 ? f.utilidad/f.ventas : 0;
                const barW = maxVentas>0 ? (f.ventas/maxVentas*100).toFixed(1) : 0;
                html += `<tr class="lc-child ${rid}-child" style="display:${expanded?'table-row':'none'}">
                    <td class="fab-rank-num" data-label="#" style="text-align:right">${idx+1}</td>
                    <td data-label="Fabricante">${f.Fabricante}</td>
                    <td class="r fab-bar-cell" data-label="Ventas $"><span class="fab-bar-bg" style="width:${barW}%;background:${qc.bg}22"></span><b>${fF(f.ventas)}</b></td>
                    <td class="r" data-label="Margen"><span class="pill ${mg>=.40?'hi':'mi'}">${fP(mg)}</span></td>
                    <td class="r" data-label="% Participación">${fP(f.pct)}</td>
                </tr>`;
            });
        });
        tbody.innerHTML = html;
        document.getElementById('fab-chart-wrap').style.height = Math.max(220, paretoList.length*34 + 60) + 'px';
        const maxVal = Math.max(...paretoList.map(f => f.ventas)) || 0;
        charts['fabricantes'] = new Chart(document.getElementById('chart-fabricantes'), {
            type: 'bar',
            data: {
                labels: paretoList.map(f => f.Fabricante),
                datasets: [{
                    data: paretoList.map(f => f.ventas),
                    backgroundColor: paretoList.map((f,i) => QUART_COLORS[f.cuartil].bg),
                    borderRadius: 5,
                    barThickness: 16
                }]
            },
            options: {
                indexAxis: 'y', responsive: true, maintainAspectRatio: false,
                layout: { padding: { right: 34 } },
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        ...PREMIUM_TOOLTIP_OPTS,
                        callbacks: {
                            label: ctx => {
                                const f  = paretoList[ctx.dataIndex];
                                const mg = f.ventas>0 ? f.utilidad/f.ventas : 0;
                                return [' Ventas: ' + fF(f.ventas), ' Margen: ' + fP(mg), ' % Participación: ' + fP(f.pct)];
                            }
                        }
                    },
                    datalabels: {
                        display: true, anchor: 'end', align: 'end',
                        color: '#2B3A42', font: { weight: '600', size: 9.5 },
                        formatter: (value, ctx) => fP(paretoList[ctx.dataIndex].pct)
                    }
                },
                scales: {
                    x: { grid:{display:false}, border:{display:false}, suggestedMax:maxVal*1.15, ticks:{font:{size:9},color:'#666666',callback:v=>fM(v)} },
                    y: { grid:{display:false}, border:{display:false}, ticks:{font:{size:10,weight:'600'},color:'#16232B'} }
                }
            }
        });
        fitTables();
    }
    // ── PESTAÑAS "Resumen <Mes>" (mes(es) cerrados detectados en VentasMesCurso) ──
    function updateResumenAnterior(slug){
        const info = RESUMENES_ANT[slug];
        if(!info) return;
        const data  = info.data.filter(r => active.has(r.NombreSucursal));
        const dates = [...new Set(data.map(r => r.FechaStr))].sort();
        const byDate = aggByDate(data);
        const tbody = document.getElementById('tabla-resumen-' + slug);
        if(!tbody) return;
        const valid = dates.filter(d => byDate[d]);
        if(!valid.length){
            tbody.innerHTML = '<tr><td colspan="7" class="empty">Selecciona sucursales para ver el detalle.</td></tr>';
            dc('resumen-'+slug+'-ventas');
            dc('resumen-'+slug+'-tickets');
            fitTables();
            return;
        }
        let tUn=0,tV=0,tU=0,tTk=0;
        const rows = valid.map(d => {
            const r  = byDate[d];
            const mg = r.ventas>0 ? r.utilidad/r.ventas : 0;
            const pill = mg>=.40 ? 'hi' : 'mi';
            const dn = DAYS[new Date(d+'T12:00:00').getDay()];
            const mAbr = MESES_ABR[parseInt(d.slice(5,7),10)-1];
            tUn+=r.unidades; tV+=r.ventas; tU+=r.utilidad; tTk+=r.tickets;
            return `<tr><td class="date" data-label="Fecha">${d.slice(8)} ${mAbr}</td><td class="dayname" data-label="Día">${dn}</td><td class="r" data-label="Unidades">${fN(r.unidades)}</td><td class="r" data-label="Ventas $"><b>${fF(r.ventas)}</b></td><td class="r" data-label="Utilidad">${fF(r.utilidad)}</td><td class="r" data-label="Margen"><span class="pill ${pill}">${fP(mg)}</span></td><td class="r" data-label="Tickets"><b>${fN(r.tickets)}</b></td></tr>`;
        }).join('');
        const totMg = tV>0 ? tU/tV : 0;
        tbody.innerHTML = rows + `<tr class="total-row"><td data-label="" colspan="2"><b>TOTAL ${info.label.toUpperCase()} ${info.anio}</b></td><td class="r" data-label="Unidades">${fN(tUn)}</td><td class="r" data-label="Ventas $">${fF(tV)}</td><td class="r" data-label="Utilidad">${fF(tU)}</td><td class="r" data-label="Margen"><span class="pill ${totMg>=.40?'hi':'mi'}">${fP(totMg)}</span></td><td class="r" data-label="Tickets">${fN(tTk)}</td></tr>`;
        const activeSucs = getActiveSucs();
        const multi      = activeSucs.length > 1;
        updateChartGeneric('resumen-'+slug+'-ventas',  'ventas',  data, valid, byDate, activeSucs, multi);
        updateChartGeneric('resumen-'+slug+'-tickets', 'tickets', data, valid, byDate, activeSucs, multi);
        fitTables();
    }
    buildButtons();
    buildMesButtons();
    refreshMesesDisponibilidad();
    refreshSucursalesDisponibilidad();
    document.getElementById('filter-bar-mes').style.display = 'none';
    document.getElementById('filter-bar-suc').style.display = 'none';
    updateObjetivos();
    setTimeout(fitTables, 100);
});
</script>
</body>
</html>"""
    html = html_template.replace("__FECHA_VALOR__",   fecha_reporte)
    html = html.replace("__FECHA_INFO__",             fecha_info)
    html = html.replace("__MES_HEADER__",             mes_header)
    html = html.replace("__DATA_JSON__",              data_json)
    html = html.replace("__LINEA_JSON__",             linea_json)
    html = html.replace("__HISTORICO_JSON__",         historico_json)
    html = html.replace("__TOP_ART_JSON__",           top_art_json)
    html = html.replace("__LINEAS_CAT_JSON__",        lineas_cat_json)
    html = html.replace("__FABRICANTES_JSON__",       fabricantes_json)
    html = html.replace("__PRESUPUESTO_JSON__",       presupuesto_json)
    html = html.replace("__CATEGORIAS_DIARIO_JSON__", categorias_diario_json)
    html = html.replace("__PERIODOS_JSON__",          periodos_json)
    html = html.replace("__SUCURSALES_JSON__",        sucursales_json)
    html = html.replace("__CURRENT_PERIOD_JSON__",    current_period_json)
    html = html.replace("__TABS_NAV_EXTRA__",         tabs_nav_extra)
    html = html.replace("__TABS_CONTENT_EXTRA__",     tabs_content_extra)
    html = html.replace("__RESUMENES_ANT_JSON__",     resumenes_ant_json)
    html = html.replace("__LOGO_B64__",               logo_b64)
    return html
def main():
    try:
        print(f"Leyendo: {EXCEL_PATH}")
        xl  = pd.read_excel(EXCEL_PATH, sheet_name=None)
        vmc = xl["VentasMesCurso"].copy()
        vm  = xl["VentasMensuales"].copy()
        tkt = xl["TicketsMensuales"].copy()
        dim = xl["TablasDimensión"].copy()
        objetivos_df = xl.get("ObjetivosVentas")
        if objetivos_df is None:
            print("⚠️  No se encontró la hoja 'ObjetivosVentas'; el pronóstico se generará sin presupuestos (quedarán en blanco).")
            objetivos_df = pd.DataFrame(columns=["Sucursal", "Presupuesto"])
        else:
            objetivos_df = objetivos_df.copy()
        suc = dim[["Clave sucursal","Nombre de sucursal"]].dropna().drop_duplicates()
        suc.columns = ["ClaveSucursal","NombreSucursal"]
        suc["ClaveSucursal"] = pd.to_numeric(suc["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
        # Los segmentadores de sucursal (botones) se ordenan por FechaApertura
        # (de más antigua a más reciente); las sucursales sin esa fecha se
        # colocan al final, ordenadas por ClaveSucursal (ID) ascendente.
        lista_sucursales = ordenar_sucursales_por_apertura(suc, objetivos_df)
        art     = dim[["Artículo","Línea"]].dropna().drop_duplicates("Artículo")
        art_dim = dim[["Artículo","Línea","Categoría","Descripción","Fabricante"]].dropna(subset=["Artículo"]).drop_duplicates("Artículo")
        print("Procesando datos...")
        # ── Separar VentasMesCurso: a inicios de mes puede traer mezcladas
        # ventas del mes que ya cerró junto con las del mes en curso. ──
        vmc_actual, meses_anteriores = separar_mes_actual_anterior(vmc, FECHA_BASE)
        agg = procesar_mes_curso(vmc_actual, suc, BOL_EXCLUIR)  # SOLO mes en curso real
        resumenes_anteriores = []
        for info in meses_anteriores:
            agg_ant = procesar_mes_curso(info["df"], suc, BOL_EXCLUIR)
            resumenes_anteriores.append({
                "label": info["label"], "anio": info["anio"], "slug": info["slug"], "agg": agg_ant
            })
        linea_agg       = procesar_lineas(vm, art, suc, BOL_EXCLUIR)
        historico_agg   = procesar_historico(vm, tkt, suc, BOL_EXCLUIR)
        top_art_agg     = procesar_top_articulos(vm, art_dim, suc, BOL_EXCLUIR)
        lineas_cat_agg  = procesar_lineas_categoria(vm, art_dim, suc, BOL_EXCLUIR)
        fabricantes_agg = procesar_fabricantes(vm, art_dim, suc, BOL_EXCLUIR)
        # El pronóstico vs. presupuesto usa 'agg', que ya sólo trae el mes en
        # curso real (p.ej. agosto), por lo que días transcurridos/operativos
        # y el pronóstico de cierre no se contaminan con ventas de julio.
        presupuesto_agg = procesar_presupuesto(agg, objetivos_df, suc, FECHA_BASE)
        # Ventas diarias por Línea/Categoría del mes en curso, para graficar
        # la tendencia día a día en 'Resumen Mes Actual'.
        categorias_diario_agg = procesar_categorias_diario(vmc_actual, art_dim, suc, BOL_EXCLUIR)
        print("Generando HTML final...")
        fecha_reporte, fecha_info, mes_header = formatear_fechas(FECHA_BASE)
        current_period_label = periodo_label_actual(FECHA_BASE)
        html = generar_html(agg, linea_agg, historico_agg, top_art_agg, lineas_cat_agg, fabricantes_agg,
                            presupuesto_agg, categorias_diario_agg, lista_sucursales, fecha_reporte, fecha_info, mes_header,
                            current_period_label, resumenes_anteriores)
        Path(OUTPUT_PATH).write_text(html, encoding="utf-8")
        print(f"✅ Dashboard generado exitosamente en: {OUTPUT_PATH}")
    except Exception as e:
        import traceback
        print(f"❌ Error crítico durante la ejecución: {e}")
        traceback.print_exc()
if __name__ == "__main__":
    main()