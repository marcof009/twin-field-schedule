#!/usr/bin/env python3
"""
generate_data.py
Lee los archivos de /data/ y genera /docs/data.json
Correr localmente o via GitHub Actions cuando cambia algo en /data/
"""

import geopandas as gpd
import pandas as pd
import json, os, re

BASE     = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, 'data')
OUT_DIR  = os.path.join(BASE, 'docs')
os.makedirs(OUT_DIR, exist_ok=True)

# ── 1. SHAPEFILES ──────────────────────────────────────────────────────────
print("Leyendo shapefiles...")
lotes = []
SHP_MAP = [
    ('la_manga_poli',   'fina_la_manga',   'La Manga'),
    ('la_turca_poli',   'fina_la_turca',   'La Turca'),
    ('san_fermin_poli', 'fina_san_fermin', 'San Fermín'),
    ('becutti_poli',    'fina_becutti',    'Becutti'),
]
for shp_name, layer, campo_label in SHP_MAP:
    path = os.path.join(DATA_DIR, f"{shp_name}.shp")
    if not os.path.exists(path):
        print(f"  AVISO: {shp_name}.shp no encontrado"); continue
    gdf = gpd.read_file(path).to_crs("EPSG:4326")
    for _, row in gdf.iterrows():
        geom = row.geometry
        rings = []
        if geom.geom_type == 'Polygon':
            rings.append([[round(c[0],5), round(c[1],5)] for c in geom.exterior.coords])
        elif geom.geom_type == 'MultiPolygon':
            for poly in geom.geoms:
                rings.append([[round(c[0],5), round(c[1],5)] for c in poly.exterior.coords])
        lotes.append({
            'name':  str(row['Name']),   'layer': layer,
            'campo': campo_label,        'area':  str(row['area']),
            'r26':   str(row.get('recurso26','') or ''),
            'r25':   str(row.get('recurso25','') or ''),
            'r24':   str(row.get('recurso24','') or ''),
            'r23':   str(row.get('recurso23','') or ''),
            'rings': rings,
        })
    print(f"  {campo_label}: {len(gdf)} lotes")

# ── 2. PULVERIZACIONES ─────────────────────────────────────────────────────
# Formato: planilla maestra con una fila por aplicación
# Columnas fijas: 0=Fecha, 1=Campo, 2=Lote, 3=Cultivo, 4=Has, 5=Estado,
#                 6=Rec Nro, 7=Observacion — desde col 14 en adelante: productos
print("Leyendo pulverizaciones...")
PULV_FILE = os.path.join(DATA_DIR, 'PULVERIZACIONES_La_Alicia.xlsx')
CAMPO_MAP = {
    'LA MANGA':   'fina_la_manga',
    'LA TURCA':   'fina_la_turca',
    'LA TURCA ':  'fina_la_turca',
    'SAN FERMIN': 'fina_san_fermin',
    'BECUTTI':    'fina_becutti',
}
pulv = {}
if os.path.exists(PULV_FILE):
    df = pd.read_excel(PULV_FILE, sheet_name='LA ALICIA', header=None, engine='openpyxl')
    headers = df.iloc[4]
    # Columnas de productos (col 14 en adelante con nombre)
    product_cols = {
        i: str(headers[i]).strip()
        for i in range(14, len(headers))
        if pd.notna(headers[i]) and str(headers[i]).strip() not in ('', 'nan')
    }
    for _, row in df.iloc[7:].iterrows():
        raw_campo = str(row[1]).strip().upper() if pd.notna(row[1]) else ''
        campo_key = CAMPO_MAP.get(raw_campo)
        if not campo_key: continue
        lote = str(row[2]).strip() if pd.notna(row[2]) else ''
        if not lote or lote == 'nan': continue
        lote_key = ('LOTE ' + lote.upper()) if not lote.upper().startswith('LOTE') else lote.upper()
        # Fecha
        fecha = ''
        if pd.notna(row[0]):
            try:    fecha = pd.to_datetime(row[0]).strftime('%d/%m/%Y')
            except: fecha = str(row[0])
        # Productos usados (dosis != 0/vacío)
        productos = [
            {'n': product_cols[i], 'd': str(row[i]).strip()}
            for i in product_cols
            if i < len(row) and pd.notna(row[i])
            and str(row[i]).strip() not in ('', 'nan', '0')
        ]
        if not productos: continue
        def sv(v): return (str(v).strip() if pd.notna(v) else '').replace('nan', '')
        pulv.setdefault(campo_key, {}).setdefault(lote_key, []).append({
            'f': fecha, 'c': sv(row[3]), 'e': sv(row[5]),
            'r': sv(row[6]), 'o': sv(row[7]), 'p': productos,
        })
    for campo, lotes_p in pulv.items():
        n = sum(len(v) for v in lotes_p.values())
        print(f"  {campo}: {len(lotes_p)} lotes, {n} aplicaciones")
else:
    print(f"  AVISO: {PULV_FILE} no encontrado")

# ── 3. ANÁLISIS DE SUELOS ──────────────────────────────────────────────────
print("Leyendo análisis de suelos...")
SUELOS_FILE = os.path.join(DATA_DIR, 'analisis_suelos.xlsx')
PARAMS = ['N-NO3','Fósforo','N Min.','S-Sulfato','MO','Zinc','pH','Humedad','Boro']
suelos = {}
if os.path.exists(SUELOS_FILE):
    df_s = pd.read_excel(SUELOS_FILE, sheet_name='Datos Completos', header=1)
    for _, row in df_s.iterrows():
        campo = str(row['Campo']).strip() if pd.notna(row['Campo']) else ''
        lote  = str(row['Lote']).strip()  if pd.notna(row['Lote'])  else ''
        if not campo or not lote: continue
        ck = 'fina_' + campo.lower().replace(' ', '_')
        suelos.setdefault(ck, {}).setdefault(lote, {})
        anio = str(int(row['Año'])) if pd.notna(row['Año']) else '?'
        if anio not in suelos[ck][lote]:
            suelos[ck][lote][anio] = {
                'archivo':  str(row['Archivo']) if pd.notna(row['Archivo']) else '',
                'perfiles': []
            }
        perfil = {'prof': str(row['Prof.']) if pd.notna(row['Prof.']) else ''}
        for col in PARAMS:
            if col in df_s.columns and pd.notna(row[col]):
                perfil[col] = float(row[col])
        suelos[ck][lote][anio]['perfiles'].append(perfil)
    total = sum(len(l) for l in suelos.values())
    print(f"  {total} lotes con análisis")
else:
    print(f"  AVISO: {SUELOS_FILE} no encontrado")

# ── 4. SIEMBRA ─────────────────────────────────────────────────────────────
print("Leyendo siembra...")
SIEMBRA_FILE = os.path.join(DATA_DIR, 'siembra.xlsx')
siembra = {}
CAMPO_KEY_MAP = {
    'LA MANGA':   'fina_la_manga',
    'LA TURCA':   'fina_la_turca',
    'SAN FERMIN': 'fina_san_fermin',
    'BECUTTI':    'fina_becutti',
}
if os.path.exists(SIEMBRA_FILE):
    df_si = pd.read_excel(SIEMBRA_FILE, sheet_name=0, header=0)
def sf(v, d=0):
            try:
                f = float(v)
                return d if f != f else f
            except:
                return d
    for _, row in df_si.iterrows():
        campo = str(row.get('Campo','')).strip().upper()
        lote  = str(row.get('Lote','')).strip()
        if not campo or not lote: continue
        ck = CAMPO_KEY_MAP.get(campo)
        if not ck: continue
        siembra.setdefault(ck, {})[lote] = {
            'cultivo':  str(row.get('Cultivo','')),
            'variedad': str(row.get('Variedad','')),
            'kg_ha':    sf(row.get('Kg_ha')),
            'map_ha':   sf(row.get('MAP_ha')),
            'dap_ha':   sf(row.get('DAP_ha')),
            'ant':      str(row.get('Antecesor','')),
            'pg':       sf(row.get('PG')),
            'p1000':    sf(row.get('P1000')),
        }
    total = sum(len(l) for l in siembra.values())
    print(f"  {total} lotes con datos de siembra")
else:
    print(f"  AVISO: {SIEMBRA_FILE} no encontrado")

# ── 5. GUARDAR ─────────────────────────────────────────────────────────────
output = {'lotes': lotes, 'pulv': pulv, 'suelos': suelos, 'siembra': siembra}
out_path = os.path.join(OUT_DIR, 'data', 'la_alicia.json')
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, separators=(',', ':'))
size = os.path.getsize(out_path)
print(f"\nOK — data.json: {out_path} ({size:,} bytes, {size//1024} KB)")
