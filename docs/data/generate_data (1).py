#!/usr/bin/env python3
"""
generate_data.py
Lee los archivos de /data/ y genera /docs/data/<cliente>.json por cada cliente
"""

import geopandas as gpd
import pandas as pd
import json, os, re

BASE     = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, 'data')
OUT_DIR  = os.path.join(BASE, 'docs', 'data')
os.makedirs(OUT_DIR, exist_ok=True)

def safe_float(v, d=0):
    try:
        f = float(v)
        return d if f != f else f
    except:
        return d

def sv(v):
    return (str(v).strip() if pd.notna(v) else '').replace('nan', '')

# ── CLIENTES CONFIG ──
CLIENTES = {
    'la_alicia': {
        'shp': [
            ('la_manga_poli',   'fina_la_manga',   'La Manga'),
            ('la_turca_poli',   'fina_la_turca',   'La Turca'),
            ('san_fermin_poli', 'fina_san_fermin', 'San Fermín'),
            ('becutti_poli',    'fina_becutti',    'Becutti'),
        ],
        'pulv_file': 'PULVERIZACIONES_La_Alicia.xlsx',
        'pulv_sheet': 'LA ALICIA',
        'pulv_campo_map': {
            'LA MANGA':   'fina_la_manga',
            'LA TURCA':   'fina_la_turca',
            'LA TURCA ':  'fina_la_turca',
            'SAN FERMIN': 'fina_san_fermin',
            'BECUTTI':    'fina_becutti',
        },
        'suelos_file': 'analisis_suelos.xlsx',
        'suelos_campo_prefix': 'fina_',
    },
    'alan': {
        'shp': [
            ('da_poli',        'don_axel',  'Don Axel'),
            ('seeber_poli',    'seeber',    'Seeber'),
            ('posborg_poli',   'posborg',   'Posborg'),
            ('gasparini_poli', 'gasparini', 'Gasparini'),
            ('gonzales_poli',  'gonzales',  'Gonzales'),
            ('fuentes_poli',   'fuentes',   'Fuentes'),
            ('pedone_poli',    'pedone',    'Pedone'),
            ('sorensen_poli',  'sorensen',  'Sorensen'),
        ],
        'pulv_file': None,
        'suelos_file': 'resumen_analisis_suelos_alan.xlsx',
    },
}

# ── SIEMBRA (shared, filtered by Cliente column) ──
print("Leyendo siembra...")
SIEMBRA_FILE = os.path.join(DATA_DIR, 'siembra.xlsx')
siembra_all = {}  # {cliente: {campo_key: {lote: {...}}}}

SIEMBRA_CAMPO_MAP = {
    # La Alicia
    'LA MANGA':   ('la_alicia', 'fina_la_manga'),
    'LA TURCA':   ('la_alicia', 'fina_la_turca'),
    'SAN FERMIN': ('la_alicia', 'fina_san_fermin'),
    'BECUTTI':    ('la_alicia', 'fina_becutti'),
    # Alan
    'DON AXEL':   ('alan', 'don_axel'),
    'POSBORG':    ('alan', 'posborg'),
    'SEEBER':     ('alan', 'seeber'),
    'FUENTE':     ('alan', 'fuentes'),
    'FUENTES':    ('alan', 'fuentes'),
    'GONZALES':   ('alan', 'gonzales'),
    'GASPARINI':  ('alan', 'gasparini'),
    'PEDONE':     ('alan', 'pedone'),
    'SORENSEN':   ('alan', 'sorensen'),
}

if os.path.exists(SIEMBRA_FILE):
    df_si = pd.read_excel(SIEMBRA_FILE, sheet_name=0, header=0)
    has_cliente_col = 'Cliente' in df_si.columns
    for _, row in df_si.iterrows():
        campo_raw = sv(row.get('Campo','')).upper()
        lote_raw  = sv(row.get('Lote',''))
        if not campo_raw or not lote_raw: continue

        # Determine client and campo_key
        if has_cliente_col:
            cliente = sv(row.get('Cliente','')).lower()
            mapping = SIEMBRA_CAMPO_MAP.get(campo_raw)
            campo_key = mapping[1] if mapping else campo_raw.lower().replace(' ','_')
        else:
            mapping = SIEMBRA_CAMPO_MAP.get(campo_raw)
            if not mapping: continue
            cliente, campo_key = mapping

        lote_key = lote_raw if lote_raw.upper().startswith('LOTE') else f'LOTE {lote_raw}'
        siembra_all.setdefault(cliente, {}).setdefault(campo_key, {})[lote_key] = {
            'cultivo':  sv(row.get('Cultivo','')),
            'variedad': sv(row.get('Variedad','')),
            'kg_ha':    safe_float(row.get('Kg_ha')),
            'map_ha':   safe_float(row.get('MAP_ha')),
            'dap_ha':   safe_float(row.get('DAP_ha')),
            'ant':      sv(row.get('Antecesor','')),
            'siembra':  sv(row.get('siembra','')),
            'pg':       safe_float(row.get('PG')),
            'p1000':    safe_float(row.get('P1000')),
            'analisis': sv(row.get('Analisis','')),
        }
    for cli, campos in siembra_all.items():
        total = sum(len(l) for l in campos.values())
        print(f"  {cli}: {total} lotes")
else:
    print(f"  AVISO: {SIEMBRA_FILE} no encontrado")

# ── GENERATE JSON PER CLIENT ──
for cliente_key, config in CLIENTES.items():
    print(f"\n{'='*40}")
    print(f"Generando {cliente_key}.json...")

    # 1. Lotes from SHP
    lotes = []
    for shp_name, layer, campo_label in config['shp']:
        path = os.path.join(DATA_DIR, f"{shp_name}.shp")
        if not os.path.exists(path):
            print(f"  AVISO: {shp_name}.shp no encontrado"); continue
        gdf = gpd.read_file(path)
        if gdf.crs is None: gdf = gdf.set_crs("EPSG:4326")
        gdf = gdf.to_crs("EPSG:4326")
        name_col = 'lote' if 'lote' in gdf.columns else 'Name'
        for _, row in gdf.iterrows():
            geom = row.geometry
            rings = []
            if geom.geom_type == 'Polygon':
                rings.append([[round(c[0],5), round(c[1],5)] for c in geom.exterior.coords])
            elif geom.geom_type == 'MultiPolygon':
                for poly in geom.geoms:
                    rings.append([[round(c[0],5), round(c[1],5)] for c in poly.exterior.coords])
            lotes.append({
                'name':  str(row[name_col]),
                'layer': layer,
                'campo': campo_label,
                'area':  str(row['area']),
                'r26':   sv(row.get('recurso26','')),
                'r25':   sv(row.get('recurso25','')),
                'r24':   sv(row.get('recurso24','')),
                'r23':   sv(row.get('recurso23','')),
                'rings': rings,
            })
        print(f"  {campo_label}: {len(gdf)} lotes")

    # 2. Pulverizaciones
    pulv = {}
    pulv_file = config.get('pulv_file')
    if pulv_file:
        PULV_PATH = os.path.join(DATA_DIR, pulv_file)
        if os.path.exists(PULV_PATH):
            df = pd.read_excel(PULV_PATH, sheet_name=config.get('pulv_sheet','Sheet1'),
                               header=None, engine='openpyxl')
            headers = df.iloc[4]
            product_cols = {
                i: str(headers[i]).strip()
                for i in range(14, len(headers))
                if pd.notna(headers[i]) and str(headers[i]).strip() not in ('', 'nan')
            }
            campo_map = config.get('pulv_campo_map', {})
            for _, row in df.iloc[7:].iterrows():
                raw_campo = str(row[1]).strip().upper() if pd.notna(row[1]) else ''
                campo_key = campo_map.get(raw_campo)
                if not campo_key: continue
                lote = str(row[2]).strip() if pd.notna(row[2]) else ''
                if not lote or lote == 'nan': continue
                lote_key = ('LOTE ' + lote.upper()) if not lote.upper().startswith('LOTE') else lote.upper()
                fecha = ''
                if pd.notna(row[0]):
                    try:    fecha = pd.to_datetime(row[0]).strftime('%d/%m/%Y')
                    except: fecha = str(row[0])
                productos = [
                    {'n': product_cols[i], 'd': str(row[i]).strip()}
                    for i in product_cols
                    if i < len(row) and pd.notna(row[i])
                    and str(row[i]).strip() not in ('', 'nan', '0')
                ]
                if not productos: continue
                pulv.setdefault(campo_key, {}).setdefault(lote_key, []).append({
                    'f': fecha, 'c': sv(row[3]), 'e': sv(row[5]),
                    'r': sv(row[6]), 'o': sv(row[7]), 'p': productos,
                })
            for campo, lotes_p in pulv.items():
                n = sum(len(v) for v in lotes_p.values())
                print(f"  Pulv {campo}: {len(lotes_p)} lotes, {n} aplicaciones")

    # 3. Suelos
    suelos = {}
    suelos_file = config.get('suelos_file')
    if suelos_file:
        SUELOS_PATH = os.path.join(DATA_DIR, suelos_file)
        if os.path.exists(SUELOS_PATH):
            if cliente_key == 'la_alicia':
                PARAMS = ['N-NO3','Fósforo','N Min.','S-Sulfato','MO','Zinc','pH','Humedad','Boro']
                df_s = pd.read_excel(SUELOS_PATH, sheet_name='Datos Completos', header=1)
                for _, row in df_s.iterrows():
                    campo = sv(row.get('Campo','')); lote = sv(row.get('Lote',''))
                    if not campo or not lote: continue
                    ck = 'fina_' + campo.lower().replace(' ', '_')
                    suelos.setdefault(ck, {}).setdefault(lote, {})
                    anio = str(int(row['Año'])) if pd.notna(row.get('Año')) else '?'
                    if anio not in suelos[ck][lote]:
                        suelos[ck][lote][anio] = {'archivo': sv(row.get('Archivo','')), 'perfiles': []}
                    perfil = {'prof': sv(row.get('Prof.',''))}
                    for col in PARAMS:
                        if col in df_s.columns and pd.notna(row[col]):
                            perfil[col] = float(row[col])
                    suelos[ck][lote][anio]['perfiles'].append(perfil)
            total = sum(len(l) for l in suelos.values())
            print(f"  Suelos: {total} lotes")

    # 4. Siembra for this client
    siembra = siembra_all.get(cliente_key, {})

    # 5. Save
    output = {'lotes': lotes, 'pulv': pulv, 'suelos': suelos, 'siembra': siembra}
    raw = json.dumps(output, ensure_ascii=False, separators=(',', ':'))
    assert 'NaN' not in raw, f"NaN encontrado en {cliente_key}.json!"
    out_path = os.path.join(OUT_DIR, f'{cliente_key}.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(raw)
    size = os.path.getsize(out_path)
    print(f"  OK — {cliente_key}.json: {size:,} bytes ({size//1024} KB)")

print("\n✓ Todos los JSONs generados")
