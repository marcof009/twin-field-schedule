#!/usr/bin/env python3
"""
generate_data.py
Reads Excel files and shapefiles from /data/ and generates /docs/data.json
Run locally or via GitHub Actions.
"""

import geopandas as gpd
import pandas as pd
import json
import re
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, 'data')
OUT_DIR  = os.path.join(BASE, 'docs')
os.makedirs(OUT_DIR, exist_ok=True)

# ── 1. SHAPEFILES ──────────────────────────────────────────────────────────
print("Reading shapefiles...")
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
        print(f"  WARNING: {shp_name}.shp not found, skipping")
        continue
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
            'name':  str(row['Name']),
            'layer': layer,
            'campo': campo_label,
            'area':  str(row['area']),
            'r26':   str(row.get('recurso26', '') or ''),
            'r25':   str(row.get('recurso25', '') or ''),
            'r24':   str(row.get('recurso24', '') or ''),
            'r23':   str(row.get('recurso23', '') or ''),
            'rings': rings,
        })
    print(f"  {campo_label}: {len(gdf)} lotes")

# ── 2. PULVERIZACIONES ─────────────────────────────────────────────────────
def parse_pulv_sheet(path, sheet_name):
    df = pd.read_excel(path, sheet_name=sheet_name, header=None)
    result = {}
    cur_lote = cur_date = cur_cultivo = cur_estado = cur_rec = cur_obs = None
    cur_prods = []

    def save():
        if cur_date and cur_lote is not None and cur_prods:
            result.setdefault(cur_lote, []).append({
                'f': cur_date, 'c': cur_cultivo, 'e': cur_estado,
                'r': cur_rec,  'o': cur_obs,     'p': cur_prods[:]
            })

    for _, row in df.iterrows():
        v = [str(row[i]).strip() if pd.notna(row[i]) else ''
             for i in range(min(9, len(row)))]
        while len(v) < 9:
            v.append('')

        if v[0].upper().startswith('LOTE') and v[1] == '' and v[6] == '':
            save(); cur_date = None; cur_prods = []
            cur_lote = v[0]
            result.setdefault(cur_lote, [])
            continue

        if re.match(r'\d{2}/\d{2}/\d{4}', v[0]):
            save()
            cur_date = v[0]; cur_cultivo = v[1]
            cur_estado = v[3]; cur_rec = v[4]; cur_obs = v[5]
            cur_prods = []
            if v[6]:
                cur_prods.append({'n': v[6], 'd': v[7]})
            continue

        if v[0] == '' and v[6]:
            cur_prods.append({'n': v[6], 'd': v[7]})

    save()
    return result

print("Reading pulverizaciones...")
pulv = {}
PULV_FILE = os.path.join(DATA_DIR, 'pulverizaciones.xlsx')
PULV_SHEETS = {
    'LA MANGA':   'fina_la_manga',
    'LA TURCA':   'fina_la_turca',
    'SAN FERMIN': 'fina_san_fermin',
    'BECUTTI':    'fina_becutti',
}
if os.path.exists(PULV_FILE):
    for sheet, layer in PULV_SHEETS.items():
        try:
            data = parse_pulv_sheet(PULV_FILE, sheet)
            pulv[layer] = data
            n_apps = sum(len(v) for v in data.values())
            print(f"  {layer}: {len(data)} lotes, {n_apps} aplicaciones")
        except Exception as e:
            print(f"  WARNING: could not read sheet {sheet}: {e}")
            pulv[layer] = {}
else:
    print("  WARNING: pulverizaciones.xlsx not found")

# ── 3. ANÁLISIS DE SUELOS ──────────────────────────────────────────────────
print("Reading análisis de suelos...")
suelos = {}
SUELOS_FILE = os.path.join(DATA_DIR, 'analisis_suelos.xlsx')
PARAMS = ['N-NO3','Fósforo','N Min.','S-Sulfato','MO','Zinc','pH','Humedad','Boro']

if os.path.exists(SUELOS_FILE):
    try:
        df_s = pd.read_excel(SUELOS_FILE, sheet_name='Datos Completos', header=1)
        for _, row in df_s.iterrows():
            campo = str(row['Campo']).strip() if pd.notna(row['Campo']) else ''
            lote  = str(row['Lote']).strip()  if pd.notna(row['Lote'])  else ''
            if not campo or not lote:
                continue
            campo_key = 'fina_' + campo.lower().replace(' ', '_')
            suelos.setdefault(campo_key, {}).setdefault(lote, {})
            anio = str(int(row['Año'])) if pd.notna(row['Año']) else '?'
            if anio not in suelos[campo_key][lote]:
                suelos[campo_key][lote][anio] = {
                    'archivo':  str(row['Archivo']) if pd.notna(row['Archivo']) else '',
                    'perfiles': []
                }
            perfil = {'prof': str(row['Prof.']) if pd.notna(row['Prof.']) else ''}
            for col in PARAMS:
                if col in df_s.columns and pd.notna(row[col]):
                    perfil[col] = float(row[col])
            suelos[campo_key][lote][anio]['perfiles'].append(perfil)
        total = sum(len(l) for l in suelos.values())
        print(f"  {total} lotes con análisis")
    except Exception as e:
        print(f"  WARNING: could not read analisis_suelos.xlsx: {e}")
else:
    print("  WARNING: analisis_suelos.xlsx not found")

# ── 4. SIEMBRA ─────────────────────────────────────────────────────────────
print("Reading siembra...")
siembra = {}
SIEMBRA_FILE = os.path.join(DATA_DIR, 'siembra.xlsx')

if os.path.exists(SIEMBRA_FILE):
    try:
        df_si = pd.read_excel(SIEMBRA_FILE, sheet_name=0, header=0)
        # Expected columns: Campo, Lote, Cultivo, Variedad, Kg_ha, MAP_ha, DAP_ha,
        #                   Antecesor, PG, P1000, Analisis
        CAMPO_KEY_MAP = {
            'LA MANGA':   'fina_la_manga',
            'LA TURCA':   'fina_la_turca',
            'SAN FERMIN': 'fina_san_fermin',
            'BECUTTI':    'fina_becutti',
        }
        for _, row in df_si.iterrows():
            campo = str(row.get('Campo','')).strip().upper()
            lote  = str(row.get('Lote','')).strip()
            if not campo or not lote:
                continue
            ck = CAMPO_KEY_MAP.get(campo)
            if not ck:
                continue
            siembra.setdefault(ck, {})[lote] = {
                'cultivo':  str(row.get('Cultivo','')),
                'variedad': str(row.get('Variedad','')),
                'kg_ha':    float(row.get('Kg_ha', 0) or 0),
                'map_ha':   float(row.get('MAP_ha', 0) or 0),
                'dap_ha':   float(row.get('DAP_ha', 0) or 0),
                'ant':      str(row.get('Antecesor','')),
                'pg':       float(row.get('PG', 0) or 0),
                'p1000':    float(row.get('P1000', 0) or 0),
                'analisis': str(row.get('Analisis','')),
            }
        total = sum(len(l) for l in siembra.values())
        print(f"  {total} lotes con datos de siembra")
    except Exception as e:
        print(f"  WARNING: could not read siembra.xlsx: {e}")
else:
    print("  WARNING: siembra.xlsx not found — skipping")

# ── 5. WRITE OUTPUT ────────────────────────────────────────────────────────
output = {
    'lotes':   lotes,
    'pulv':    pulv,
    'suelos':  suelos,
    'siembra': siembra,
}
out_path = os.path.join(OUT_DIR, 'data.json')
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, separators=(',', ':'))

size = os.path.getsize(out_path)
print(f"\nOK — data.json written to {out_path} ({size:,} bytes)")
EOF