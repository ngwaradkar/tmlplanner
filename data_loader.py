import pandas as pd
import numpy as np
import io
import re
import os
import openpyxl
import pyxlsb
import sqlite3
import streamlit as st
import requests

def _hash_upload_buffer(b):
    """
    Custom hash for st.cache_data: our upload buffers are io.BytesIO objects
    with a .name attribute set (to mimic Streamlit's UploadedFile). Streamlit's
    default hasher treats anything with a .name as a real file handle and
    calls os.path.getmtime() on it -- which crashes, since these are in-memory
    buffers, not files that exist at that path. Hash by actual content instead.
    """
    pos = b.tell()
    b.seek(0)
    data = b.read()
    b.seek(pos)
    return data

# Cache TTL for parsed-file results: long enough that clicking around the
# dashboard (Telegram send, tab switches, filters) doesn't re-parse the same
# Excel/xlsb file from scratch every single rerun, short enough to self-heal
# if a cache entry is ever wrong.
_CACHE_TTL_SECONDS = 1800

def clean_part_number(val):
    if pd.isna(val):
        return None
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    # If the string is empty or just whitespace or '0', return None or '0'
    if not val_str or val_str.isspace():
        return None
    return val_str

def _detect_html_content(filepath_or_buffer):
    """
    Detects whether a file (given as an on-disk path string OR an in-memory
    file-like buffer, e.g. a Streamlit UploadedFile returned by
    st.file_uploader) is actually an HTML table disguised with an .xls
    extension -- a common export format from plant reporting systems (PPC,
    DPT Plan, etc). This works identically for both cases, so uploads made
    through the browser file uploader are detected the same way as files
    picked up automatically from disk.

    Returns (is_html: bool, html_content: str or None). Buffers are left at
    position 0 afterwards so downstream code can still read them.
    """
    name = filepath_or_buffer if isinstance(filepath_or_buffer, str) else getattr(filepath_or_buffer, 'name', '')
    ext = os.path.splitext(str(name).lower())[1]

    # Plant reporting systems (PPC, DPT, etc.) frequently export HTML tables mislabeled as .xls, .xlsx, etc.
    if ext not in ['.xls', '.xlsx', '.xlsb', '.html', '.htm', '']:
        return False, None

    try:
        if isinstance(filepath_or_buffer, str):
            with open(filepath_or_buffer, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
        else:
            filepath_or_buffer.seek(0)
            raw = filepath_or_buffer.read()
            filepath_or_buffer.seek(0)
            content = raw.decode('utf-8', errors='ignore') if isinstance(raw, bytes) else raw
    except Exception:
        return False, None

    first_chars = content[:500].lower()
    if '<html' in first_chars or '<style' in first_chars or '<table' in first_chars:
        # Some plant systems (confirmed on Shop Wise Report exports) stamp every
        # <td> with rowspan="0" colspan="0" as generator boilerplate. Per the
        # HTML5 spec, rowspan/colspan="0" means "span to the end of the table/row" --
        # so parsers (incl. pandas.read_html) collapse every row after the first
        # into NaN, since row 0's cells are treated as spanning the whole table.
        # These files don't actually want spanning; normalize to "1" so each
        # cell is read as its own row/column.
        content = re.sub(r'(rowspan|colspan)\s*=\s*"0"', r'\1="1"', content, flags=re.IGNORECASE)
        content = re.sub(r"(rowspan|colspan)\s*=\s*'0'", r"\1='1'", content, flags=re.IGNORECASE)
        return True, content
    return False, None

@st.cache_data(show_spinner=False, ttl=_CACHE_TTL_SECONDS, hash_funcs={io.BytesIO: _hash_upload_buffer})
def load_bom(filepath_or_buffer):
    """
    Loads BOM details.xlsx and returns a cleaned DataFrame.
    Expected columns: 'Short Vehicle Code', 'Front Wiring', 'Cockpit ', 'Engine'
    """
    if not isinstance(filepath_or_buffer, str):
        try:
            filepath_or_buffer.seek(0)
        except Exception:
            pass
    # openpyxl engine is used for .xlsx
    df = pd.read_excel(filepath_or_buffer, engine='openpyxl')
    
    # Strip whitespace from headers
    df.columns = [str(c).strip() for c in df.columns]
    
    # Required columns check
    req = ['Short Vehicle Code', 'Front Wiring', 'Cockpit', 'Engine']
    for r in req:
        if r not in df.columns:
            # Check with partial match
            matched_col = [c for c in df.columns if r.lower() in c.lower()]
            if matched_col:
                df.rename(columns={matched_col[0]: r}, inplace=True)
            else:
                raise ValueError(f"Required BOM column '{r}' not found. Available: {list(df.columns)}")
                
    # Keep only target columns and clean data
    df = df[['Short Vehicle Code', 'Front Wiring', 'Cockpit', 'Engine']].copy()
    df.dropna(subset=['Short Vehicle Code'], inplace=True)
    
    # Clean part numbers and vehicle codes
    df['Short Vehicle Code'] = df['Short Vehicle Code'].astype(str).str.strip()
    df['Front Wiring'] = df['Front Wiring'].apply(clean_part_number)
    df['Cockpit'] = df['Cockpit'].apply(clean_part_number)
    df['Engine'] = df['Engine'].apply(clean_part_number)
    
    # Deduplicate by Short Vehicle Code
    df.drop_duplicates(subset=['Short Vehicle Code'], keep='first', inplace=True)
    return df

@st.cache_data(show_spinner=False, ttl=_CACHE_TTL_SECONDS, hash_funcs={io.BytesIO: _hash_upload_buffer})
def load_float_report(filepath_or_buffer):
    """
    Loads PPC Float Report (which can be .xlsb or .xls disguised HTML) and returns a cleaned DataFrame.
    Required columns: 'BIW NUMBER', 'VEHICLE CODE', 'SHOP', 'PBS LIFT', 'HOLD BY', 'REASONS S', 'VIN'
    """
    if not isinstance(filepath_or_buffer, str):
        try:
            filepath_or_buffer.seek(0)
        except Exception:
            pass
    # Check if the file is an HTML table disguised as .xls (works for both
    # on-disk paths and in-memory uploads from st.file_uploader)
    is_html, html_content = _detect_html_content(filepath_or_buffer)

    if is_html:
        dfs = pd.read_html(io.StringIO(html_content))
        if not dfs:
            raise ValueError("No tables found in PPC Float Report HTML file.")
        df = dfs[0]
        # Promote first row to headers
        df.columns = [str(c).strip() for c in df.iloc[0]]
        df = df.iloc[1:].reset_index(drop=True)
    else:
        # Read the file using pd.read_excel
        if isinstance(filepath_or_buffer, str):
            ext = os.path.splitext(filepath_or_buffer.lower())[1]
            if ext == '.xlsb':
                df = pd.read_excel(filepath_or_buffer, engine='pyxlsb')
            else:
                df = pd.read_excel(filepath_or_buffer)
        else:
            # For stream or bytes, let's try to detect/read
            try:
                df = pd.read_excel(filepath_or_buffer, engine='pyxlsb')
            except Exception:
                try:
                    filepath_or_buffer.seek(0)
                    df = pd.read_excel(filepath_or_buffer)
                except Exception:
                    try:
                        filepath_or_buffer.seek(0)
                        dfs = pd.read_html(filepath_or_buffer)
                        df = dfs[0]
                        df.columns = [str(c).strip() for c in df.iloc[0]]
                        df = df.iloc[1:].reset_index(drop=True)
                    except Exception as e:
                        raise ValueError(f"Failed to load float report: {e}")

    df.columns = [str(c).strip() for c in df.columns]
    
    # Check columns
    req = ['BIW NUMBER', 'VEHICLE CODE', 'SHOP', 'PBS LIFT', 'HOLD BY', 'REASONS S', 'VIN']
    for r in req:
        if r not in df.columns:
            matched_col = [c for c in df.columns if r.lower() in c.lower() or (r == 'VEHICLE CODE' and c.lower() in ['vc', 'vehicle_code', 'vehicle code'])]
            if matched_col:
                df.rename(columns={matched_col[0]: r}, inplace=True)
                
    # Parse dates
    for date_col in ['BIW LIFTING', 'PTCED', 'SEALANT', 'TOPCOAT', 'PBS LIFT']:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], format='mixed', dayfirst=True, errors='coerce')
            
    # Clean BIW NUMBER and VEHICLE CODE
    if 'BIW NUMBER' in df.columns:
        df['BIW NUMBER'] = df['BIW NUMBER'].apply(lambda x: str(int(x)) if pd.notna(x) and str(x).strip().replace('.0','').isdigit() else str(x).strip())
    if 'VEHICLE CODE' in df.columns:
        df['VEHICLE CODE'] = df['VEHICLE CODE'].astype(str).str.strip()
        df['VC'] = df['VEHICLE CODE']
    elif 'VC' in df.columns:
        df['VC'] = df['VC'].astype(str).str.strip()
        df['VEHICLE CODE'] = df['VC']
    else:
        df['VEHICLE CODE'] = ""
        df['VC'] = ""
    
    # Group by BIW NUMBER to aggregate multiple hold reasons
    # If a cab appears multiple times, it is because of multiple hold entries.
    # We combine 'HOLD BY' and 'REASONS S' by joining with a comma.
    def join_non_null(series):
        non_nulls = series.dropna().astype(str).str.strip()
        non_nulls = [n for n in non_nulls if n and n != 'nan']
        return ', '.join(dict.fromkeys(non_nulls)) # Deduplicated join
        
    def first_non_null(series):
        non_nulls = series.dropna()
        return non_nulls.iloc[0] if not non_nulls.empty else None
        
    agg_funcs = {}
    for col in df.columns:
        if col == 'BIW NUMBER':
            continue
        elif col in ['HOLD BY', 'REASONS S']:
            agg_funcs[col] = join_non_null
        else:
            agg_funcs[col] = first_non_null
            
    df_dedup = df.groupby('BIW NUMBER').agg(agg_funcs).reset_index()
    
    # Clean empty strings resulting from join_non_null back to None
    df_dedup['HOLD BY'] = df_dedup['HOLD BY'].replace('', None)
    df_dedup['REASONS S'] = df_dedup['REASONS S'].replace('', None)
    
    return df_dedup

@st.cache_data(show_spinner=False, ttl=_CACHE_TTL_SECONDS, hash_funcs={io.BytesIO: _hash_upload_buffer})
def load_paint_summary_report(filepath_or_buffer):
    """
    Loads PPC Float Report Paint Summary (HTML or Excel format) and returns aggregated model stage totals.
    Uses dynamic header scanning and flexible row parsing to handle formatting variations across .xls and .xlsb files.
    """
    if not isinstance(filepath_or_buffer, str):
        try:
            filepath_or_buffer.seek(0)
        except Exception:
            pass
            
    is_html, html_content = _detect_html_content(filepath_or_buffer)

    if is_html:
        dfs = pd.read_html(io.StringIO(html_content))
        if not dfs:
            return {}
        df = dfs[0]
    else:
        if isinstance(filepath_or_buffer, str):
            ext = os.path.splitext(filepath_or_buffer.lower())[1]
            if ext == '.xlsb':
                df = pd.read_excel(filepath_or_buffer, engine='pyxlsb')
            else:
                df = pd.read_excel(filepath_or_buffer)
        else:
            try:
                df = pd.read_excel(filepath_or_buffer, engine='pyxlsb')
            except Exception:
                try:
                    filepath_or_buffer.seek(0)
                    df = pd.read_excel(filepath_or_buffer)
                except Exception:
                    try:
                        filepath_or_buffer.seek(0)
                        dfs = pd.read_html(filepath_or_buffer)
                        if not dfs:
                            return {}
                        df = dfs[0]
                    except Exception as e:
                        raise ValueError(f"Failed to load Paint Float Summary report: {e}")
                
    pf_map = {
        'HORNBILL': 'PUNCH',
        'NOVA': 'PUNCH.EV',
        'ETURNA': 'HARRIER.EV',
        'GRAVITAS': 'SAFARI',
        'Q5': 'HARRIER',
        'TAYRONA': 'SAFARI.EV'
    }
    
    stage_cols = [
        'TOTAL FLOAT', 'PBS FLOAT', 'PBS TO POLISHING', 'POLISHING TO TOPCOAT',
        'TOPCOAT TO WETSANDING G ROOFBLACK', 'TOPCOAT TO WETSANDING G FRESH',
        'WETSANDING G TO SEALANT', 'TOTAL UPTO SEALANT', 'PT ENTRY TO SEALANT',
        'BIW LIFTING G TO PT', 'PT BYPASS'
    ]
    
    col_map = {}
    header_row_idx = None
    for r_idx in range(min(10, len(df))):
        row_vals = [str(x).upper().strip() for x in df.iloc[r_idx].values]
        if any('TOTAL FLOAT' in x for x in row_vals):
            header_row_idx = r_idx
            for c_idx, val in enumerate(row_vals):
                if 'TOTAL FLOAT' in val and 'TOTAL FLOAT' not in col_map:
                    col_map['TOTAL FLOAT'] = c_idx
                elif 'PBS FLOAT' in val:
                    col_map['PBS FLOAT'] = c_idx
                elif 'PBS TO POLISHING' in val:
                    col_map['PBS TO POLISHING'] = c_idx
                elif 'POLISHING TO TOPCOAT' in val:
                    col_map['POLISHING TO TOPCOAT'] = c_idx
                elif 'TOPCOAT TO WETSANDING' in val or 'TOPCOAT TO WET' in val:
                    if 'ROOF' in val or 'BLACK' in val:
                        col_map['TOPCOAT TO WETSANDING G ROOFBLACK'] = c_idx
                    else:
                        col_map['TOPCOAT TO WETSANDING G FRESH'] = c_idx
                elif 'WETSANDING' in val and 'SEAL' in val:
                    col_map['WETSANDING G TO SEALANT'] = c_idx
                elif 'UPTO SEALANT' in val or 'UPTO SEALENT' in val:
                    col_map['TOTAL UPTO SEALANT'] = c_idx
                elif 'PT ENTRY' in val:
                    col_map['PT ENTRY TO SEALANT'] = c_idx
                elif 'BIW LIFTING' in val:
                    col_map['BIW LIFTING G TO PT'] = c_idx
                elif 'PT BYPASS' in val:
                    col_map['PT BYPASS'] = c_idx
            break
            
    # Fallback to default indices (5..15) if header parsing didn't map everything
    default_indices = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]
    for stage_name, default_idx in zip(stage_cols, default_indices):
        if stage_name not in col_map:
            col_map[stage_name] = default_idx
            
    model_totals = {m: {c: 0 for c in stage_cols} for m in ['PUNCH', 'PUNCH.EV', 'HARRIER.EV', 'SAFARI', 'HARRIER', 'SAFARI.EV']}
    
    start_r = (header_row_idx + 1) if header_row_idx is not None else 0
    for idx in range(start_r, len(df)):
        row = df.iloc[idx]
        cell_texts = [str(row.iloc[i]).strip() for i in range(min(4, len(row))) if pd.notna(row.iloc[i])]
        full_text = ' '.join(cell_texts).upper()
        
        for pf_name, model in pf_map.items():
            if pf_name in full_text and 'TOTAL' in full_text and 'GRAND' not in full_text and 'SUB' not in full_text:
                for stage_name in stage_cols:
                    c_idx = col_map[stage_name]
                    if c_idx < len(row):
                        try:
                            v_str = str(row.iloc[c_idx]).strip()
                            v = int(float(v_str)) if v_str and v_str.lower() != 'nan' else 0
                        except Exception:
                            v = 0
                        model_totals[model][stage_name] += v
                break

    return model_totals

@st.cache_data(show_spinner=False, ttl=_CACHE_TTL_SECONDS, hash_funcs={io.BytesIO: _hash_upload_buffer})
def load_vgl(filepath_or_buffer):
    """
    Loads Vehicle Generation List / DPT Plan VIN Generation Report (.xls, .xlsx, .xlsb, HTML).
    Supports both old VGL format (cab-by-cab) and new DPT Plan format (VC-grouped VIN counts).
    """
    if not isinstance(filepath_or_buffer, str):
        try:
            filepath_or_buffer.seek(0)
        except Exception:
            pass

    is_html, html_content = _detect_html_content(filepath_or_buffer)

    pf_map = {
        'HORNBILL': 'PUNCH',
        'NOVA': 'PUNCH.EV',
        'ETURNA': 'HARRIER.EV',
        'GRAVITAS': 'SAFARI',
        'Q5': 'HARRIER',
        'TAYRONA': 'SAFARI.EV'
    }

    if is_html and html_content:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')
        rows = soup.find_all('tr')
        if rows:
            header = [td.get_text(strip=True) for td in rows[0].find_all(['td', 'th'])]
            header_str = ' '.join(header).upper()
            
            # Check if this is new DPT Plan VIN generation format (has MARKET and ProductFamily headers)
            if ('PRODUCTFAMILY' in header_str or 'PRODUCT FAMILY' in header_str or 'MARKET' in header_str) and 'SR NO' not in header_str:
                parsed_rows = []
                current_pf = ''
                for r in rows[1:]:
                    tds = [td.get_text(strip=True) for td in r.find_all(['td', 'th'])]
                    if len(tds) < 4:
                        continue
                    market, pf, vc, sales_desc = tds[0], tds[1], tds[2], tds[3]
                    vin_str = tds[5] if len(tds) > 5 else (tds[4] if len(tds) > 4 else '0')
                    
                    if pf:
                        current_pf = pf.strip().upper()
                        
                    if not vc or vc.upper() in ['TOTAL', 'GRAND TOTAL'] or market.upper() in ['TOTAL', 'GRAND TOTAL']:
                        continue
                        
                    try:
                        vin_cnt = int(float(str(vin_str).strip()))
                    except Exception:
                        vin_cnt = 0
                        
                    model = pf_map.get(current_pf, 'UNKNOWN')
                    if model == 'UNKNOWN':
                        for k, v in pf_map.items():
                            if k in sales_desc.upper() or k in current_pf:
                                model = v
                                break
                                
                    parsed_rows.append({
                        'VEHICLE CODE': str(vc).strip(),
                        'ProductFamily': current_pf,
                        'SALES DESCRIPTION': sales_desc,
                        'Model': model,
                        'Model_Family': model,
                        'VIN_Count': vin_cnt,
                        'CREATED DATE': pd.Timestamp.now()
                    })
                df_res = pd.DataFrame(parsed_rows)
                if not df_res.empty:
                    return df_res

    # Fallback reading via pd.read_html / pd.read_excel
    if is_html:
        dfs = pd.read_html(io.StringIO(html_content))
        df = dfs[0]
        df.columns = [str(c).strip() for c in df.iloc[0]]
        df = df.iloc[1:].reset_index(drop=True)
    else:
        if isinstance(filepath_or_buffer, str):
            ext = os.path.splitext(filepath_or_buffer.lower())[1]
            if ext == '.xlsb':
                df = pd.read_excel(filepath_or_buffer, engine='pyxlsb')
            else:
                df = pd.read_excel(filepath_or_buffer)
        else:
            fname = getattr(filepath_or_buffer, 'name', '').lower()
            if fname.endswith('.xlsb'):
                try:
                    df = pd.read_excel(filepath_or_buffer, engine='pyxlsb')
                except Exception:
                    filepath_or_buffer.seek(0)
                    df = pd.read_excel(filepath_or_buffer)
            else:
                try:
                    df = pd.read_excel(filepath_or_buffer)
                except Exception:
                    try:
                        filepath_or_buffer.seek(0)
                        df = pd.read_excel(filepath_or_buffer, engine='pyxlsb')
                    except Exception:
                        filepath_or_buffer.seek(0)
                        dfs = pd.read_html(filepath_or_buffer)
                        df = dfs[0]
                        df.columns = [str(c).strip() for c in df.iloc[0]]
                        df = df.iloc[1:].reset_index(drop=True)
        
    df.columns = [str(c).strip() for c in df.columns]

    standard_renames = {
        'VEHICLE CODE': 'VEHICLE CODE',
        'VC': 'VEHICLE CODE',
        'VEHICLE_CODE': 'VEHICLE CODE',
        'SHORT VEHICLE CODE': 'VEHICLE CODE',
        'FULL VC': 'VEHICLE CODE',
        'BIW NUMBER': 'BIW NUMBER',
        'VIN NUMBER': 'VIN NUMBER',
        'CREATED DATE': 'CREATED DATE',
        'SHIFT': 'SHIFT'
    }
    for col in list(df.columns):
        col_clean = str(col).strip()
        for std_key, std_val in standard_renames.items():
            if col_clean.lower() == std_key.lower():
                df.rename(columns={col: std_val}, inplace=True)
                break
                
    if 'VEHICLE CODE' in df.columns:
        df['VEHICLE CODE'] = df['VEHICLE CODE'].astype(str).str.strip()
        df['VC'] = df['VEHICLE CODE']
    else:
        df['VEHICLE CODE'] = ""
        df['VC'] = ""

    # Filter out total / summary rows
    vc_mask = df['VEHICLE CODE'].str.upper().isin(['TOTAL', 'GRAND TOTAL', 'NAN', ''])
    if 'MARKET' in df.columns:
        vc_mask = vc_mask | df['MARKET'].astype(str).str.strip().str.upper().isin(['TOTAL', 'GRAND TOTAL'])
    df = df[~vc_mask].reset_index(drop=True)

    if 'ProductFamily' in df.columns:
        df['ProductFamily'] = df['ProductFamily'].replace(r'^\s*$', None, regex=True).ffill()
    if 'PRODUCT' in df.columns:
        df['PRODUCT'] = df['PRODUCT'].replace(r'^\s*$', None, regex=True).ffill()

    vin_col = None
    for c in df.columns:
        c_str = str(c).strip().upper()
        if c_str in ['TCF/-VIN', 'TCF2-VIN', 'TCF-VIN', 'TCF1-VIN', 'TCF1_VIN', 'TCF2_VIN', 'VIN_COUNT', 'TODAY VIN', 'VIN GENERATION', 'VIN QTY', 'VIN_QTY']:
            vin_col = c
            break
        if 'VIN' in c_str and 'CODE' not in c_str and 'NUMBER' not in c_str and 'DESC' not in c_str and 'PLAN' not in c_str:
            vin_col = c
            break

    if vin_col:
        df['VIN_Count'] = pd.to_numeric(df[vin_col], errors='coerce').fillna(0).astype(int)
    else:
        df['VIN_Count'] = 1

    if 'BIW NUMBER' in df.columns:
        df['BIW NUMBER'] = df['BIW NUMBER'].apply(lambda x: str(int(float(x))) if pd.notna(x) and str(x).strip().replace('.0','').replace('e+','').replace('+','').isdigit() else str(x).strip())
    
    def map_model_family(row):
        pf = str(row.get('PRODUCT', row.get('ProductFamily', ''))).strip().upper()
        sd = str(row.get('SALES DESCRIPTION', row.get('SALES DESC', ''))).strip().upper()
        vc = str(row.get('VEHICLE CODE', row.get('VC', ''))).strip().upper()
        
        if 'NOVA' in pf or 'PUNCH.EV' in sd or 'PUNCH EV' in sd:
            return 'PUNCH.EV'
        elif 'HORNBILL' in pf or 'PUNCH' in sd:
            return 'PUNCH'
        elif 'ETURNA' in pf or 'HARRIER.EV' in sd or 'HARRIER EV' in sd:
            return 'HARRIER.EV'
        elif 'TAYRONA' in pf or 'SAFARI.EV' in sd or 'SAFARI EV' in sd or '54831927A' in vc:
            return 'SAFARI EV'
        elif 'GRAVITAS' in pf or 'SAFARI' in sd:
            return 'SAFARI'
        elif 'Q5' in pf or 'HARRIER' in sd:
            return 'HARRIER'
        for k, v in pf_map.items():
            if k in pf or k in sd:
                return v
        return 'UNKNOWN'
        
    df['Model_Family'] = df.apply(map_model_family, axis=1)
    df['Model'] = df['Model_Family']
    
    if 'CREATED DATE' in df.columns:
        df['CREATED DATE'] = pd.to_datetime(df['CREATED DATE'], dayfirst=True, errors='coerce')
        
    engine_col = None
    for col in df.columns:
        c_str = str(col).strip().lower()
        if 'engine part' in c_str or 'engine no' in c_str:
            engine_col = col
            break
            
    if engine_col:
        df.rename(columns={engine_col: 'RAW_ENGINE_PART'}, inplace=True)
    else:
        df['RAW_ENGINE_PART'] = None
        
    def extract_engine(val):
        if pd.isna(val):
            return None
        val_str = str(val).split('/')[0].strip()
        if val_str.lower() in ['na', 'none', '']:
            return None
        return val_str
        
    df['Engine_Part'] = df['RAW_ENGINE_PART'].apply(extract_engine)
    
    return df

@st.cache_data(show_spinner=False, ttl=_CACHE_TTL_SECONDS, hash_funcs={io.BytesIO: _hash_upload_buffer})
def load_stock_grouped(filepath_or_buffer, sheet_name, vc_col_idx, part_col_idx, qty_col_idx, skip_rows=2):
    """
    Generic grouped-row stock parser.
    Skips skip_rows, reads columns by index, forward-fills part_number and qty,
    deduplicated to return a dict {part_number: qty}.
    """
    if not isinstance(filepath_or_buffer, str):
        try:
            filepath_or_buffer.seek(0)
        except Exception:
            pass
    # Read the sheet with no header to get index-based access
    # Detect file type
    if isinstance(filepath_or_buffer, str):
        ext = os.path.splitext(filepath_or_buffer.lower())[1]
        if ext == '.xlsb':
            df = pd.read_excel(filepath_or_buffer, sheet_name=sheet_name, engine='pyxlsb', header=None)
        else:
            df = pd.read_excel(filepath_or_buffer, sheet_name=sheet_name, engine='xlrd' if ext == '.xls' else 'openpyxl', header=None)
    else:
        # Uploaded bytes - let's check name if possible, otherwise try xlrd/openpyxl
        # We try to read it as excel
        try:
            df = pd.read_excel(filepath_or_buffer, sheet_name=sheet_name, engine='openpyxl', header=None)
        except Exception:
            try:
                # Seek back to 0
                filepath_or_buffer.seek(0)
                df = pd.read_excel(filepath_or_buffer, sheet_name=sheet_name, engine='xlrd', header=None)
            except Exception:
                filepath_or_buffer.seek(0)
                df = pd.read_excel(filepath_or_buffer, sheet_name=sheet_name, engine='pyxlsb', header=None)
                
    # Slice the data, skipping headers
    data = df.iloc[skip_rows:].copy()
    
    # Grab the target columns
    vc_col = data.iloc[:, vc_col_idx]
    part_col = data.iloc[:, part_col_idx]
    qty_col = data.iloc[:, qty_col_idx]
    
    # Clean part numbers and quantity
    parts_clean = part_col.apply(clean_part_number)
    
    # Build forward-filled list
    current_part = None
    current_qty = None
    
    part_to_qty = {}
    vc_to_part = {}
    
    for idx, (vc_val, part_val, qty_val) in enumerate(zip(vc_col, parts_clean, qty_col)):
        if part_val is not None:
            current_part = part_val
            # Clean qty value
            if pd.isna(qty_val) or str(qty_val).strip() == '':
                current_qty = 0
            else:
                try:
                    # If it is a string representing a float/int, parse it
                    qty_str = str(qty_val).strip()
                    if '.' in qty_str:
                        current_qty = int(float(qty_str))
                    else:
                        current_qty = int(qty_str)
                except ValueError:
                    current_qty = 0
                    
        if current_part is not None:
            part_to_qty[current_part] = current_qty
            
        if pd.notna(vc_val):
            vc_str = str(vc_val).strip()
            # Clean floats in VCs
            if vc_str.endswith('.0'):
                vc_str = vc_str[:-2]
            if len(vc_str) >= 8 and vc_str[0].isdigit():
                vc_to_part[vc_str] = current_part
                
    return part_to_qty, vc_to_part

def classify_files(uploaded_files):
    """
    Classifies a list of uploaded file wrappers by their filenames.
    Returns a dict mapping classification category to the file object.
    """
    classifications = {}
    
    for f in uploaded_files:
        name = f.name.lower()
        # Top priority: Shop_Wise report
        if name.startswith('shop_wise') or name.startswith('shopwise') or name.startswith('shop-wise') or 'shop_wise' in name or 'shopwise' in name or 'shop-wise' in name or ('shop' in name and 'report' in name):
            classifications['SHOP_WISE_REPORT'] = f
        elif 'bom' in name:
            classifications['BOM'] = f
        elif 'float' in name:
            if 'paint' in name:
                classifications['FLOAT_PAINT_SUMMARY'] = f
            else:
                classifications['FLOAT_REPORT'] = f
        elif any(k in name for k in ['vgl', 'vehicle_generation', 'generation_list', 'dpt-plan', 'dpt_plan', 'vin_generation']):
            if 'tcf2' in name or 'tcf_2' in name:
                classifications['TCF2_VGL'] = f
            else:
                classifications['TCF1_VGL'] = f
        elif 'cockpit' in name:
            if 'tcf2' in name or 'tcf-2' in name or 'tcf_2' in name or 'harrier' in name or 'safari' in name:
                classifications['TCF2_COCKPIT_STOCK'] = f
            elif 'nova' in name:
                classifications['TCF1_NOVA_COCKPIT_STOCK'] = f
            else:
                classifications['TCF1_ALTROZ_COCKPIT_STOCK'] = f
        elif 'wiring' in name or 'harness' in name:
            if 'tcf2' in name or 'tcf-2' in name or 'tcf_2' in name:
                classifications['TCF2_WIRING_STOCK'] = f
            else:
                classifications['TCF1_WIRING_STOCK'] = f
        elif 'hourly' in name:
            classifications['HOURLY_PRODUCTION'] = f
                
    return classifications

def _extract_date_tuple(filename):
    """Extracts (year, month, day) from filename like Shop_Wise_Report_08_05_2026."""
    m = re.search(r'(\d{2})[_\-](\d{2})[_\-](\d{4})', filename)
    if m:
        try:
            return (int(m.group(3)), int(m.group(1)), int(m.group(2)))
        except Exception:
            pass
    return (0, 0, 0)

def detect_and_classify_files(directory_path):
    """
    Scans directory_path (including subdirectories like 'Float reports/') and classifies all xls, xlsx, xlsb files by their names.
    Returns a dict mapping categories to their absolute file paths.
    """
    if not os.path.exists(directory_path) or not os.path.isdir(directory_path):
        return {}
        
    classifications = {}
    file_entries = []
    
    try:
        # First scan files directly in directory_path
        for f in os.listdir(directory_path):
            p = os.path.join(directory_path, f)
            if os.path.isfile(p):
                file_entries.append((f, p))
                
        # Then scan subdirectories (e.g., Float reports, TEST)
        for f_dir in os.listdir(directory_path):
            p_dir = os.path.join(directory_path, f_dir)
            if os.path.isdir(p_dir) and not f_dir.startswith('.') and f_dir != '__pycache__':
                for root, dirs, files in os.walk(p_dir):
                    for f in files:
                        file_entries.append((f, os.path.join(root, f)))
    except Exception:
        return {}
        
    for name, path in file_entries:
        ext = os.path.splitext(name.lower())[1]
        if ext not in ['.xlsx', '.xls', '.xlsb', '.csv', '.xlsm']:
            continue
            
        name_lower = name.lower()
        
        def set_cat(cat, p):
            if cat not in classifications:
                classifications[cat] = p
            else:
                existing = classifications[cat]
                try:
                    p_name = os.path.basename(p)
                    e_name = os.path.basename(existing)
                    p_date = _extract_date_tuple(p_name)
                    e_date = _extract_date_tuple(e_name)
                    
                    p_test = 'test' in p.lower()
                    e_test = 'test' in existing.lower()
                    
                    if e_test and not p_test:
                        classifications[cat] = p
                    elif not e_test and p_test:
                        pass
                    elif p_date > e_date:
                        classifications[cat] = p
                    elif p_date == e_date and os.path.getmtime(p) > os.path.getmtime(existing):
                        classifications[cat] = p
                except Exception:
                    pass

        # Top priority: Shop_Wise report
        if name_lower.startswith('shop_wise') or name_lower.startswith('shopwise') or name_lower.startswith('shop-wise') or 'shop_wise' in name_lower or 'shopwise' in name_lower or 'shop-wise' in name_lower or ('shop' in name_lower and 'report' in name_lower):
            set_cat('SHOP_WISE_REPORT', path)
        elif 'bom' in name_lower:
            set_cat('BOM', path)
        elif 'float' in name_lower:
            if 'paint' in name_lower:
                set_cat('FLOAT_PAINT_SUMMARY', path)
            else:
                set_cat('FLOAT_REPORT', path)
        elif any(k in name_lower for k in ['vgl', 'vehicle_generation', 'generation_list', 'dpt-plan', 'dpt_plan', 'vin_generation']):
            if 'tcf2' in name_lower or 'tcf_2' in name_lower:
                set_cat('TCF2_VGL', path)
            else:
                set_cat('TCF1_VGL', path)
        elif 'cockpit' in name_lower:
            if 'tcf2' in name_lower or 'tcf-2' in name_lower or 'tcf_2' in name_lower or 'harrier' in name_lower or 'safari' in name_lower:
                set_cat('TCF2_COCKPIT_STOCK', path)
            elif 'nova' in name_lower:
                set_cat('TCF1_NOVA_COCKPIT_STOCK', path)
            else:
                set_cat('TCF1_ALTROZ_COCKPIT_STOCK', path)
        elif 'wiring' in name_lower or 'harness' in name_lower:
            if 'tcf2' in name_lower or 'tcf-2' in name_lower or 'tcf_2' in name_lower:
                set_cat('TCF2_WIRING_STOCK', path)
            else:
                set_cat('TCF1_WIRING_STOCK', path)
        elif 'hourly' in name_lower:
            set_cat('HOURLY_PRODUCTION', path)
        elif 'dashboard files' in name_lower:
            if 'HOURLY_PRODUCTION' not in classifications:
                set_cat('HOURLY_PRODUCTION', path)

    # If master Dashboard files.xlsm exists, unpack and prioritize its sheets for all 6 categories
    for name, path in file_entries:
        name_lower = name.lower()
        if 'dashboard files' in name_lower and (name_lower.endswith('.xlsm') or name_lower.endswith('.xlsx')):
            try:
                with open(path, 'rb') as f_db:
                    db_bytes = f_db.read()
                db_buffers = parse_onedrive_workbook(db_bytes)
                for cat_key, buf_val in db_buffers.items():
                    if buf_val is not None:
                        classifications[cat_key] = buf_val
                break
            except Exception:
                pass
                
    return classifications

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "clear_to_build.db")

def init_db():
    """Initializes the database and table if they do not exist.

    Also self-heals a pre-existing bom_details table that's missing the
    short_vehicle_code PRIMARY KEY constraint (a table in that state can
    result from the old save_bom_to_db(), which used to_sql(if_exists='replace')
    and silently dropped the constraint on every full BOM re-upload). Without
    this repair, save_single_bom_entry()'s ON CONFLICT upsert would fail on
    any database created before this fix.
    """
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bom_details (
            short_vehicle_code TEXT PRIMARY KEY,
            front_wiring TEXT,
            cockpit TEXT,
            engine TEXT
        )
    """)
    conn.commit()

    # Check whether short_vehicle_code is actually a PRIMARY KEY on the table
    # that exists right now (it may predate this schema and lack it).
    cursor.execute("PRAGMA table_info(bom_details)")
    cols_info = cursor.fetchall()  # (cid, name, type, notnull, dflt_value, pk)
    has_pk = any(c[1] == 'short_vehicle_code' and c[5] > 0 for c in cols_info)

    if not has_pk:
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS bom_details_fixed (
                short_vehicle_code TEXT PRIMARY KEY,
                front_wiring TEXT,
                cockpit TEXT,
                engine TEXT
            )
        """)
        # Keep the last row per short_vehicle_code if the broken table has duplicates
        cursor.execute("""
            INSERT OR REPLACE INTO bom_details_fixed (short_vehicle_code, front_wiring, cockpit, engine)
            SELECT short_vehicle_code, front_wiring, cockpit, engine FROM bom_details
        """)
        cursor.execute("DROP TABLE bom_details")
        cursor.execute("ALTER TABLE bom_details_fixed RENAME TO bom_details")
        conn.commit()

    conn.close()

def save_bom_to_db(df):
    """Overwrites the database table with the provided BOM DataFrame.

    Uses DELETE + INSERT into the existing table (rather than pandas'
    to_sql(if_exists='replace'), which drops and recreates the table using a
    default schema and silently discards the short_vehicle_code PRIMARY KEY
    constraint defined in init_db() -- that constraint is required for
    save_single_bom_entry()'s upsert to work).
    """
    if df is None or df.empty:
        return
    init_db()
    conn = sqlite3.connect(DB_PATH)
    db_df = df.copy()
    db_df.columns = ['short_vehicle_code', 'front_wiring', 'cockpit', 'engine']
    cursor = conn.cursor()
    cursor.execute("DELETE FROM bom_details")
    db_df.to_sql('bom_details', conn, if_exists='append', index=False)
    conn.commit()
    conn.close()

def save_single_bom_entry(short_vehicle_code, front_wiring, cockpit, engine):
    """
    Upserts one BOM row (by Short Vehicle Code) into the database, without
    touching any other rows. Used by the homepage 'missing BOM' entry form so
    filling in one short VC never risks wiping the rest of the BOM table
    (unlike save_bom_to_db, which replaces the whole table).
    """
    short_vehicle_code = str(short_vehicle_code).strip()
    if not short_vehicle_code:
        raise ValueError("Short Vehicle Code is required.")
    front_wiring = clean_part_number(front_wiring)
    cockpit = clean_part_number(cockpit)
    engine = clean_part_number(engine)
    init_db()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO bom_details (short_vehicle_code, front_wiring, cockpit, engine)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(short_vehicle_code) DO UPDATE SET
            front_wiring = excluded.front_wiring,
            cockpit = excluded.cockpit,
            engine = excluded.engine
    """, (short_vehicle_code, front_wiring, cockpit, engine))
    conn.commit()
    conn.close()

def load_bom_from_db():
    """Loads the BOM data from the database. Returns None if empty or table does not exist."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM bom_details")
        count = cursor.fetchone()[0]
        if count == 0:
            return None
        df = pd.read_sql("SELECT * FROM bom_details", conn)
        # Rename back to the standard application columns
        df.rename(columns={
            'short_vehicle_code': 'Short Vehicle Code',
            'front_wiring': 'Front Wiring',
            'cockpit': 'Cockpit',
            'engine': 'Engine'
        }, inplace=True)
        return df
    except Exception:
        return None
    finally:
        conn.close()

def init_engine_db():
    """Initializes the engine stocks table if it does not exist."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS engine_stocks (
            tcf_line TEXT,
            engine_part_no TEXT PRIMARY KEY,
            model TEXT,
            ta_code TEXT,
            clearance_qty INTEGER
        )
    """)
    conn.commit()
    conn.close()

def save_engine_stocks_to_db(df):
    """Saves the edited engine stocks DataFrame to the DB."""
    if df is None or df.empty:
        return
    init_engine_db()
    conn = sqlite3.connect(DB_PATH)
    db_df = df.copy()
    db_df.columns = ['tcf_line', 'engine_part_no', 'model', 'ta_code', 'clearance_qty']
    db_df.to_sql('engine_stocks', conn, if_exists='replace', index=False)
    conn.commit()
    conn.close()

def load_engine_stocks_from_db():
    """Loads engine stocks from the DB. Returns None if empty or table does not exist."""
    init_engine_db()
    conn = sqlite3.connect(DB_PATH)
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM engine_stocks")
        count = cursor.fetchone()[0]
        if count == 0:
            return None
        df = pd.read_sql("SELECT * FROM engine_stocks", conn)
        df.columns = ["TCF Line", "Engine Part No", "Model", "TA Code", "Clearance After 6:30AM"]
        return df
    except Exception:
        return None
    finally:
        conn.close()

def save_nova_stocks_to_db(df):
    """Saves the Punch EV Nova material stock DataFrame to the DB."""
    if df is None or df.empty:
        return
    init_db()
    conn = sqlite3.connect(DB_PATH)
    try:
        df.to_sql('nova_stocks', conn, if_exists='replace', index=False)
    except Exception:
        pass
    finally:
        conn.close()

def load_nova_stocks_from_db():
    """Loads Punch EV Nova material stocks from the DB."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql("SELECT * FROM nova_stocks", conn)
        if not df.empty:
            return df
        return None
    except Exception:
        return None
    finally:
        conn.close()

def save_model_shortages_to_db(df):
    """Saves Model-Wise Shortage DataFrame to the DB."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    try:
        if df is None or df.empty:
            conn.execute("DROP TABLE IF EXISTS model_shortages")
        else:
            df.to_sql('model_shortages', conn, if_exists='replace', index=False)
    except Exception:
        pass
    finally:
        conn.close()

def load_model_shortages_from_db():
    """Loads Model-Wise Shortages from the DB."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql("SELECT * FROM model_shortages", conn)
        if not df.empty:
            return df
        return None
    except Exception:
        return None
    finally:
        conn.close()

def init_metadata_db():
    """Initializes the metadata table if it does not exist."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS metadata (key TEXT PRIMARY KEY, value TEXT)")
    conn.commit()
    conn.close()

def save_metadata(key, value):
    """Saves a metadata key-value pair to the database."""
    init_metadata_db()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO metadata (key, value) VALUES (?, ?)", (key, str(value)))
    conn.commit()
    conn.close()

def load_metadata(key, default=None):
    """Loads a metadata value by key from the database. Returns default if not found."""
    init_metadata_db()
    conn = sqlite3.connect(DB_PATH)
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM metadata WHERE key = ?", (key,))
        row = cursor.fetchone()
        if row:
            return row[0]
        return default
    except Exception:
        return default
    finally:
        conn.close()

@st.cache_data(show_spinner=False, ttl=_CACHE_TTL_SECONDS, hash_funcs={io.BytesIO: _hash_upload_buffer})
def load_paint_summary_by_vc(filepath_or_buffer):
    """
    Loads PPC Float Report Paint Summary and returns stage float counts aggregated by Short Vehicle Code (SHORT VC).
    Returns dict: { '54734327A': {'TOTAL FLOAT': int, 'PBS FLOAT': int, 'TOTAL UPTO SEALANT': int}, ... }
    """
    if not filepath_or_buffer:
        return {}

    if not isinstance(filepath_or_buffer, str):
        try:
            filepath_or_buffer.seek(0)
        except Exception:
            pass
            
    is_html, html_content = _detect_html_content(filepath_or_buffer)

    try:
        if is_html:
            dfs = pd.read_html(io.StringIO(html_content))
            if not dfs:
                return {}
            df = dfs[0]
        else:
            if isinstance(filepath_or_buffer, str):
                ext = os.path.splitext(filepath_or_buffer.lower())[1]
                if ext == '.xlsb':
                    df = pd.read_excel(filepath_or_buffer, engine='pyxlsb')
                else:
                    df = pd.read_excel(filepath_or_buffer)
            else:
                try:
                    df = pd.read_excel(filepath_or_buffer, engine='pyxlsb')
                except Exception:
                    try:
                        filepath_or_buffer.seek(0)
                        df = pd.read_excel(filepath_or_buffer)
                    except Exception:
                        try:
                            filepath_or_buffer.seek(0)
                            dfs = pd.read_html(filepath_or_buffer)
                            if not dfs:
                                return {}
                            df = dfs[0]
                        except Exception:
                            return {}
    except Exception:
        return {}

    header_row_idx = None
    col_map = {}
    
    for r_idx in range(min(15, len(df))):
        row_vals = [str(x).upper().strip() for x in df.iloc[r_idx].values]
        if any('SHORT VC' in x or 'VEHICLE CODE' in x or x == 'VC' for x in row_vals) or any('TOTAL FLOAT' in x for x in row_vals):
            header_row_idx = r_idx
            for c_idx, val in enumerate(row_vals):
                if 'SHORT VC' in val or 'VEHICLE CODE' in val or val == 'VC':
                    col_map['VC'] = c_idx
                elif 'TOTAL FLOAT' in val and 'TOTAL FLOAT' not in col_map:
                    col_map['TOTAL FLOAT'] = c_idx
                elif 'PBS FLOAT' in val:
                    col_map['PBS FLOAT'] = c_idx
                elif 'UPTO SEALANT' in val or 'UPTO SEALENT' in val:
                    col_map['TOTAL UPTO SEALANT'] = c_idx
                elif 'PBS TO POLISHING' in val:
                    col_map['PBS TO POLISHING'] = c_idx
                elif 'POLISHING TO TOPCOAT' in val:
                    col_map['POLISHING TO TOPCOAT'] = c_idx
                elif 'TOPCOAT TO WETSANDING' in val or 'TOPCOAT TO WET' in val:
                    if 'ROOF' in val or 'BLACK' in val:
                        col_map['TOPCOAT TO WETSANDING G ROOFBLACK'] = c_idx
                    else:
                        col_map['TOPCOAT TO WETSANDING G FRESH'] = c_idx
                elif 'WETSANDING' in val and 'SEAL' in val:
                    col_map['WETSANDING G TO SEALANT'] = c_idx
            break

    vc_counts = {}
    if header_row_idx is not None and 'VC' in col_map:
        vc_cidx = col_map['VC']
        tot_cidx = col_map.get('TOTAL FLOAT')
        pbs_cidx = col_map.get('PBS FLOAT')
        seal_cidx = col_map.get('TOTAL UPTO SEALANT')
        
        start_r = header_row_idx + 1
        for idx in range(start_r, len(df)):
            row = df.iloc[idx]
            if vc_cidx < len(row):
                raw_vc = str(row.iloc[vc_cidx]).strip()
                if raw_vc and raw_vc.lower() != 'nan' and len(raw_vc) >= 8 and not raw_vc.upper().startswith('TOTAL'):
                    svc = raw_vc[:9]
                    
                    def get_int(c_idx):
                        if c_idx is not None and c_idx < len(row):
                            try:
                                v_str = str(row.iloc[c_idx]).strip()
                                return int(float(v_str)) if v_str and v_str.lower() != 'nan' else 0
                            except Exception:
                                return 0
                        return 0
                        
                    tot = get_int(tot_cidx)
                    pbs = get_int(pbs_cidx)
                    
                    if seal_cidx is not None:
                        seal = get_int(seal_cidx)
                    else:
                        seal = pbs + get_int(col_map.get('PBS TO POLISHING')) + get_int(col_map.get('POLISHING TO TOPCOAT')) + get_int(col_map.get('TOPCOAT TO WETSANDING G ROOFBLACK')) + get_int(col_map.get('TOPCOAT TO WETSANDING G FRESH')) + get_int(col_map.get('WETSANDING G TO SEALANT'))
                        
                    if svc not in vc_counts:
                        vc_counts[svc] = {'TOTAL FLOAT': 0, 'PBS FLOAT': 0, 'TOTAL UPTO SEALANT': 0}
                    vc_counts[svc]['TOTAL FLOAT'] += tot
                    vc_counts[svc]['PBS FLOAT'] += pbs
                    vc_counts[svc]['TOTAL UPTO SEALANT'] += seal
                    
    return vc_counts


# ----------------- TELEGRAM DISPATCHER UTILITY -----------------
def send_telegram_message(bot_token, chat_id, message_text):
    """
    Sends a message via Telegram Bot API with HTML formatting.
    Converts plain text lines starting with emojis/headers to HTML bold automatically.
    """
    import urllib.request
    import json

    if not bot_token or not str(bot_token).strip():
        return False, "Telegram Bot Token is missing."
    if not chat_id or not str(chat_id).strip():
        return False, "Telegram Chat ID is missing."

    token = str(bot_token).strip()
    c_id = str(chat_id).strip()
    url = f"https://api.telegram.org/bot{token}/sendMessage"

    formatted_lines = []
    for line in message_text.splitlines():
        if "<b>" in line:
            formatted_lines.append(line)
        elif line.startswith("📊") or line.startswith("🏭") or line.startswith("⚡") or line.startswith("📦") or line.startswith("⏰") or line.startswith("🚗"):
            formatted_lines.append(f"<b>{line}</b>")
        elif "• ✅" in line or "• 🚫" in line or "• 🟢" in line or "• 🟡" in line or "• 🔴" in line or "• 🚜" in line or "🚗" in line:
            parts = line.split(":", 1)
            if len(parts) == 2:
                formatted_lines.append(f"{parts[0]}: <b>{parts[1].strip()}</b>")
            else:
                formatted_lines.append(line)
        else:
            formatted_lines.append(line)
            
    final_text = "\n".join(formatted_lines)

    payload = {
        "chat_id": c_id,
        "text": final_text,
        "parse_mode": "HTML"
    }

    data = json.dumps(payload).encode('utf-8')
    req = urllib.request.Request(url, data=data, headers={'Content-Type': 'application/json'})

    try:
        with urllib.request.urlopen(req, timeout=12) as response:
            res_data = json.loads(response.read().decode('utf-8'))
            if res_data.get('ok'):
                return True, "✅ Telegram message sent successfully!"
            else:
                desc = res_data.get('description', 'Unknown error')
                return False, f"❌ Telegram API Error: {desc}"
    except Exception as e:
        return False, f"❌ HTTP/Network Error: {e}"


def send_telegram_photo(bot_token, chat_id, photo_bytes, caption="", filename="image.png"):
    """
    Sends an image file via Telegram Bot API with optional caption.
    """
    import requests

    if not bot_token or not str(bot_token).strip():
        return False, "Telegram Bot Token is missing."
    if not chat_id or not str(chat_id).strip():
        return False, "Telegram Chat ID is missing."

    token = str(bot_token).strip()
    c_id = str(chat_id).strip()
    url = f"https://api.telegram.org/bot{token}/sendPhoto"

    files = {
        'photo': (filename, photo_bytes)
    }
    data = {
        'chat_id': c_id,
        'parse_mode': 'HTML'
    }
    if caption:
        data['caption'] = caption

    try:
        res = requests.post(url, data=data, files=files, timeout=30)
        res_json = res.json()
        if res_json.get('ok'):
            return True, "✅ Image sent to Telegram successfully!"
        else:
            desc = res_json.get('description', 'Unknown error')
            return False, f"❌ Telegram API Error: {desc}"
    except Exception as e:
        return False, f"❌ HTTP/Network Error: {e}"


def send_telegram_document(bot_token, chat_id, file_bytes, caption="", filename="document.xlsx"):
    """
    Sends a document or Excel file via Telegram Bot API with optional caption.
    """
    import requests

    if not bot_token or not str(bot_token).strip():
        return False, "Telegram Bot Token is missing."
    if not chat_id or not str(chat_id).strip():
        return False, "Telegram Chat ID is missing."

    token = str(bot_token).strip()
    c_id = str(chat_id).strip()
    url = f"https://api.telegram.org/bot{token}/sendDocument"

    files = {
        'document': (filename, file_bytes)
    }
    data = {
        'chat_id': c_id,
        'parse_mode': 'HTML'
    }
    if caption:
        data['caption'] = caption

    try:
        res = requests.post(url, data=data, files=files, timeout=30)
        res_json = res.json()
        if res_json.get('ok'):
            return True, "✅ Document/Excel file sent to Telegram successfully!"
        else:
            desc = res_json.get('description', 'Unknown error')
            return False, f"❌ Telegram API Error: {desc}"
    except Exception as e:
        return False, f"❌ HTTP/Network Error: {e}"


def extract_trim_from_sales_desc(sales_desc, model_name=""):
    """Extracts standardized Trim label from Sales Description or Vehicle Code."""
    if not sales_desc or pd.isna(sales_desc):
        return "—"
    d = str(sales_desc).upper().strip()
    m = str(model_name).upper().strip()
    
    fuel = ""
    if "CNG" in d:
        fuel = " CNG"
    elif "TGDI" in d:
        fuel = " TGDI"
        
    # Trim check
    if "FEA+" in d or "FEA +" in d or "FEA PLUS" in d:
        return "FEA+" + fuel
    elif "FEA" in d or "FEARLESS" in d:
        return "FEA" + fuel
    elif "EMP+ S" in d or "EMP + S" in d:
        return "EMP+ S" + fuel
    elif "EMP" in d or "EMPOWERED" in d:
        return "EMP" + fuel
    elif "FRLR" in d:
        return "FRLR" + fuel
    elif "FRL" in d:
        return "FRL" + fuel
    elif "ACCOMP + S" in d or "ACCOMP+S" in d or "ACCOMP +S" in d:
        return "ACCOMP + S" + fuel
    elif "ACCOMP" in d or "ACC" in d:
        return "ACC" if ("HARRIER" in m or "SAFARI" in m) else ("ACCOMP" + fuel)
    elif "ADVT S" in d or "ADVT + S" in d or "ADV S" in d:
        return "ADVT S" + fuel
    elif "ADVT" in d or "ADVENTURE" in d or "ADV" in d:
        return "ADV" if ("HARRIER" in m or "SAFARI" in m) else ("ADVT" + fuel)
    elif "PURE + S" in d or "PURE+S" in d:
        return "PURE + S" + fuel
    elif "PURE +" in d or "PURE+" in d:
        return "PURE +" + fuel
    elif "PURE" in d or "PUR" in d:
        return "PUR" if ("HARRIER" in m or "SAFARI" in m) else ("PURE" + fuel)
    elif "CREATIVE" in d or "LUX" in d:
        return "CREATIVE" + fuel
    elif "SMART+" in d or "SMT+" in d:
        return "SMT+" + fuel
    elif "SMART" in d or "SMT" in d:
        return "SMT" if ("HARRIER" in m or "SAFARI" in m or "EV" in m) else ("SMART" + fuel)

    tokens = d.split()
    if len(tokens) >= 3:
        return " ".join(tokens[1:3]) + fuel
    return d + fuel


@st.cache_data(show_spinner=False, ttl=_CACHE_TTL_SECONDS, hash_funcs={io.BytesIO: _hash_upload_buffer})
def load_all_models_catalog(filepath):
    """Loads All models.xlsx spreadsheet and returns VC mapping dicts for Sales Description and Trim."""
    vc_to_desc = {}
    vc_to_trim = {}
    if not filepath or not os.path.exists(filepath):
        return vc_to_desc, vc_to_trim
        
    try:
        xl = pd.ExcelFile(filepath)
        for sheet in xl.sheet_names:
            df = pd.read_excel(filepath, sheet_name=sheet)
            desc_col = 'Sales Description' if 'Sales Description' in df.columns else None
            vc_col = 'Color VC' if 'Color VC' in df.columns else ('vc' if 'vc' in df.columns else ('SUB VC' if 'SUB VC' in df.columns else None))
            group_col = 'Group PL ' if 'Group PL ' in df.columns else ('PRODUCT' if 'PRODUCT' in df.columns else None)
            
            if desc_col and vc_col:
                for idx, row in df.iterrows():
                    full_v = str(row[vc_col]).strip()
                    s_desc = str(row[desc_col]).strip()
                    g_pl = str(row[group_col]).strip() if group_col and pd.notna(row[group_col]) else ""
                    trim_val = extract_trim_from_sales_desc(s_desc, g_pl)
                    
                    if full_v and full_v != 'nan':
                        vc_to_desc[full_v] = s_desc
                        vc_to_trim[full_v] = trim_val
                        short_v = full_v[:9]
                        vc_to_desc[short_v] = s_desc
                        vc_to_trim[short_v] = trim_val
    except Exception as e:
        print(f"Error loading All models.xlsx catalog: {e}")
        
    return vc_to_desc, vc_to_trim


def map_tcf_model_name(raw_name):
    """Maps internal code names from Shop Wise Report to clean TCF model names and filters non-TCF models."""
    name = str(raw_name).strip()
    u = name.upper()
    
    if 'HORNBILL' in u or 'PUNCH' in u:
        if 'EXP' in u:
            return 'PUNCH Exports'
        if 'EV' in u or 'NOVA' in u:
            return 'PUNCH EV'
        return 'PUNCH'
    elif 'NOVA' in u:
        return 'PUNCH EV'
    elif 'ETURNA' in u:
        return 'HARRIER EV'
    elif 'GRAVITAS' in u or 'SAFARI' in u:
        if 'EV' in u:
            return 'SAFARI EV'
        return 'SAFARI'
    elif 'Q5' in u or 'HARRIER' in u:
        if 'EV' in u or 'ETURNA' in u:
            return 'HARRIER EV'
        return 'HARRIER'
    return None


def _get_extension(filepath_or_buffer):
    if isinstance(filepath_or_buffer, str):
        return os.path.splitext(filepath_or_buffer.lower())[1]
    elif hasattr(filepath_or_buffer, 'name'):
        return os.path.splitext(str(filepath_or_buffer.name).lower())[1]
    return ''

def _reset_buffer(filepath_or_buffer):
    if hasattr(filepath_or_buffer, 'seek'):
        try:
            filepath_or_buffer.seek(0)
        except Exception:
            pass


@st.cache_data(show_spinner=False, ttl=_CACHE_TTL_SECONDS, hash_funcs={io.BytesIO: _hash_upload_buffer})
def load_shop_wise_report(filepath_or_buffer, return_debug=False):
    """
    Loads Shop Wise Production Summary Report (Shop_Wise_Report_*.xlsb, .xlsx, .xls, or HTML) and returns:
      - totals_dict: dict of plant total metrics (TCF VIN, TCF2 VIN, TCF DROP, TCF2 DROP, TCF ROLL, TCF2 ROLL, TCF IOK, TCF2 IOK, PAINT, WELD, PBS, T60, T40)
      - df_vehicles: DataFrame of TCF1 & TCF2 model-wise production breakdown with updated brand names
      - df_ta: DataFrame of Transaxle (TA) engine dispatch counts

    If return_debug=True, a 4th value (debug_info dict) is also returned. debug_info always has:
      - 'success': bool
      - 'engine_used': which read path produced the data ('html', 'pyxlsb', 'read_excel', 'xlrd', or None)
      - 'attempts': list of {'stage': str, 'error': str} for every fallback that was tried and failed
      - 'reason': short human-readable explanation when success is False
    """
    debug_info = {'success': False, 'engine_used': None, 'attempts': [], 'reason': None}

    def _fail(reason):
        debug_info['reason'] = reason
        if return_debug:
            return None, None, None, debug_info
        return None, None, None

    if not filepath_or_buffer:
        return _fail("No file was provided.")

    df = None
    try:
        _reset_buffer(filepath_or_buffer)
        
        # 1. HTML detection (common plant system exports disguised as .xls/.xlsx)
        is_html, html_content = _detect_html_content(filepath_or_buffer)
        if is_html:
            try:
                dfs = pd.read_html(io.StringIO(html_content), header=None)
                if dfs:
                    df = dfs[0]
                    debug_info['engine_used'] = 'html'
            except Exception as e:
                debug_info['attempts'].append({'stage': 'html', 'error': str(e)})

        # 2. Try pyxlsb engine if .xlsb extension
        if df is None:
            ext = _get_extension(filepath_or_buffer)
            if ext == '.xlsb':
                try:
                    df = pd.read_excel(filepath_or_buffer, sheet_name=0, engine='pyxlsb', header=None)
                    debug_info['engine_used'] = 'pyxlsb'
                except Exception as e:
                    debug_info['attempts'].append({'stage': 'pyxlsb (.xlsb ext)', 'error': str(e)})
                    _reset_buffer(filepath_or_buffer)

        # 3. Standard pandas read_excel fallback chain
        if df is None:
            try:
                df = pd.read_excel(filepath_or_buffer, sheet_name=0, header=None)
                debug_info['engine_used'] = 'read_excel'
            except Exception as e:
                debug_info['attempts'].append({'stage': 'read_excel (auto engine)', 'error': str(e)})
                _reset_buffer(filepath_or_buffer)
                try:
                    df = pd.read_excel(filepath_or_buffer, sheet_name=0, engine='pyxlsb', header=None)
                    debug_info['engine_used'] = 'pyxlsb'
                except Exception as e:
                    debug_info['attempts'].append({'stage': 'pyxlsb (fallback)', 'error': str(e)})
                    _reset_buffer(filepath_or_buffer)
                    try:
                        df = pd.read_excel(filepath_or_buffer, sheet_name=0, engine='xlrd', header=None)
                        debug_info['engine_used'] = 'xlrd'
                    except Exception as e:
                        debug_info['attempts'].append({'stage': 'xlrd', 'error': str(e)})
                        _reset_buffer(filepath_or_buffer)
                        try:
                            dfs = pd.read_html(filepath_or_buffer, header=None)
                            if dfs:
                                df = dfs[0]
                                debug_info['engine_used'] = 'html'
                        except Exception as e:
                            debug_info['attempts'].append({'stage': 'html (fallback)', 'error': str(e)})

        if df is None:
            return _fail(
                "Could not read the file with any available engine (pyxlsb / openpyxl / xlrd / HTML). "
                "See 'attempts' for the specific error from each engine tried."
            )
        if df.empty or len(df) < 2:
            return _fail(f"File was read successfully (engine: {debug_info['engine_used']}) but has too few rows ({len(df)}) to contain a report.")

        # Dynamically locate the header row (contains 'TCF', 'VIN', 'DROP', or 'DATE')
        header_row_idx = 0
        header_found = False
        for r_i in range(min(10, len(df))):
            row_vals_str = [str(x).strip().upper() for x in df.iloc[r_i].values]
            if any('TCF' in x or 'VIN' in x or 'DROP' in x or 'PAINT' in x for x in row_vals_str):
                header_row_idx = r_i
                header_found = True
                break
        if not header_found:
            debug_info['attempts'].append({
                'stage': 'header detection',
                'error': "No row in the first 10 contained 'TCF'/'VIN'/'DROP'/'PAINT'; defaulted to row 0, which may be wrong."
            })

        header_row = df.iloc[header_row_idx].values
        cols = [str(c).strip() for c in header_row]

        totals_row_idx = header_row_idx + 1
        totals_dict = {}
        if totals_row_idx < len(df):
            totals_row = df.iloc[totals_row_idx].values
            for col_idx, col_name in enumerate(cols):
                if col_idx < len(totals_row):
                    val = totals_row[col_idx]
                    try:
                        totals_dict[col_name] = int(float(str(val).strip()))
                    except Exception:
                        totals_dict[col_name] = str(val).strip()

        model_rows = []
        ta_rows = []
        data_start_idx = totals_row_idx + 1
        for i in range(data_start_idx, len(df)):
            row_vals = df.iloc[i].values
            raw_model = str(row_vals[0]).strip()
            if not raw_model or raw_model.lower() in ['nan', 'none', 'total', 'model']:
                continue

            # TA engine check
            if raw_model.isdigit():
                qty = 0
                try:
                    qty = int(float(str(row_vals[1]).strip()))
                except Exception:
                    pass
                ta_rows.append({'Transaxle (TA) Code': raw_model, 'Count': qty})
                continue

            # TCF Model check
            mapped_model = map_tcf_model_name(raw_model)
            if not mapped_model:
                continue # Skip non-TCF models like ALTROZ, NEXON, CURVV.EV

            m_dict = {'Model': mapped_model}
            for col_idx, col_name in enumerate(cols[1:], start=1):
                val = row_vals[col_idx] if col_idx < len(row_vals) else 0
                try:
                    num_v = int(float(str(val).strip()))
                    m_dict[col_name] = num_v
                except Exception:
                    m_dict[col_name] = 0
            model_rows.append(m_dict)

        simplified_model_rows = []
        for r in model_rows:
            m_name = r['Model']
            v_cnt = r.get('TCF VIN', 0) + r.get('TCF2 VIN', 0)
            d_cnt = r.get('TCF DROP', 0) + r.get('TCF2 DROP', 0)
            p_cnt = r.get('PAINT', 0)
            t60_cnt = r.get('T60', 0)
            t40_cnt = r.get('T40', 0)
            
            simplified_model_rows.append({
                'Model': m_name,
                'VIN': v_cnt,
                'Drop': d_cnt,
                'Paint Lifting': p_cnt,
                'T60': t60_cnt,
                'T40': t40_cnt
            })

        df_vehicles = pd.DataFrame(simplified_model_rows) if simplified_model_rows else None
        df_ta = pd.DataFrame(ta_rows) if ta_rows else None

        if not totals_dict and df_vehicles is None and df_ta is None:
            return _fail(
                f"File was read (engine: {debug_info['engine_used']}) and a header row was located, "
                "but no totals, model rows, or TA rows could be extracted. The sheet layout may not match "
                "what the parser expects (e.g. models not recognized in map_tcf_model_name, or shifted columns)."
            )

        # Guard against the header-detection landing on the wrong row: if it did,
        # totals_dict can still come out non-empty, just full of the wrong keys
        # (e.g. metadata labels like 'FROM DATE : ...' instead of 'TCF VIN').
        # That used to look like success even though every KPI silently reads as 0.
        core_keys = {'TCF VIN', 'TCF2 VIN', 'TCF DROP', 'TCF2 DROP', 'PAINT'}
        matched_core_keys = core_keys & set(totals_dict.keys())
        if totals_dict and not matched_core_keys:
            return _fail(
                f"File was read (engine: {debug_info['engine_used']}) but the header row detected at "
                f"index {header_row_idx} doesn't look like the real data header -- extracted keys were "
                f"{list(totals_dict.keys())} instead of expected keys like 'TCF VIN'/'TCF2 VIN'/'PAINT'. "
                "The real header row may sit outside the first 10 rows scanned, or an earlier row is "
                "being mistaken for it."
            )

        debug_info['success'] = True
        if return_debug:
            return totals_dict, df_vehicles, df_ta, debug_info
        return totals_dict, df_vehicles, df_ta
    except Exception as e:
        print(f"Error loading Shop_Wise_Report: {e}")
        return _fail(f"Unexpected error while parsing: {e}")

def download_from_onedrive(url):
    """
    Downloads the Excel file from a OneDrive/SharePoint URL, or reads from a local path.
    Returns the file content as bytes.
    """
    url = url.strip()
    
    # If it's a local file path
    if not url.startswith('http'):
        import os
        if os.path.exists(url):
            with open(url, 'rb') as f:
                return f.read()
        else:
            raise FileNotFoundError(f"Local file not found: {url}")

    # If it's a URL
    if '?' in url:
        download_url = f"{url}&download=1"
    else:
        download_url = f"{url}?download=1"
        
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }
    
    response = requests.get(download_url, headers=headers)
    response.raise_for_status()
    
    # Check if we got an HTML login page instead of the Excel file
    if 'html' in response.headers.get('Content-Type', '').lower():
        raise ValueError(
            "Received an HTML login page instead of the Excel file. "
            "This happens if the SharePoint link has 'Block Download' enabled. "
            "Please generate a link with 'Block Download' UNCHECKED, or enter the local path to the synced file (e.g. Dashboard files.xlsm)."
        )
        
    return response.content

def parse_onedrive_workbook(bytes_content):
    """
    Takes the raw bytes of the Dashboard files.xlsm, parses it, and returns 
    a dictionary of BytesIO buffers for each of the 5 required reports.
    """
    excel_file = pd.ExcelFile(io.BytesIO(bytes_content), engine='openpyxl')
    
    # Mapping of system categories to their sheet names in the workbook
    sheet_mapping = {
        'FLOAT_REPORT': 'PPC_Float_BIW_VIN',
        'FLOAT_PAINT_SUMMARY': 'PPC_Float_Paint',
        'SHOP_WISE_REPORT': 'Date_Shop_Wise',
        'TCF1_VGL': 'TCF1_VIN_Gen',
        'TCF2_VGL': 'TCF2_VIN_Gen',
        'HOURLY_PRODUCTION': 'Hourly_Production'
    }
    
    buffers = {}
    for key, sheet_name in sheet_mapping.items():
        if sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            
            buf = io.BytesIO()
            buf.name = f"{sheet_name}.xlsx"
            with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            buf.seek(0)
            buffers[key] = buf
        else:
            buffers[key] = None
            
    return buffers

def load_hourly_production(filepath_or_buffer):
    """
    Loads Hourly_Production report from Excel file, bytes buffer, or DataFrame.
    Filters to only the 4 required achievement points:
    - TCF1_VIN_GENERATION
    - TCF2_VIN_GENERATION
    - TCF1-DROP
    - TCF2-DROP
    Returns a cleaned DataFrame or None.
    """
    if filepath_or_buffer is None:
        return None
        
    try:
        if isinstance(filepath_or_buffer, pd.DataFrame):
            df = filepath_or_buffer.copy()
        elif hasattr(filepath_or_buffer, 'getvalue') or isinstance(filepath_or_buffer, (str, io.BytesIO)):
            if hasattr(filepath_or_buffer, 'seek'):
                filepath_or_buffer.seek(0)
                
            excel_file = pd.ExcelFile(filepath_or_buffer)
            sheet_target = None
            for s in excel_file.sheet_names:
                if 'hourly' in s.lower() or 'production' in s.lower():
                    sheet_target = s
                    break
            if not sheet_target:
                sheet_target = excel_file.sheet_names[0]
                
            df = pd.read_excel(excel_file, sheet_name=sheet_target)
        else:
            return None
            
        if df.empty:
            return None
            
        # Target achievement points to retain
        target_keys = {
            'TCF1_VIN_GENERATION': 'TCF1_VIN_GENERATION',
            'TCF2_VIN_GENERATION': 'TCF2_VIN_GENERATION',
            'TCF1-DROP': 'TCF1-DROP',
            'TCF2-DROP': 'TCF2-DROP'
        }
        
        def norm_key(val):
            return re.sub(r'[^A-Z0-9]', '', str(val).upper())
            
        norm_map = {norm_key(k): k for k in target_keys}
        
        # Locate achievement point column
        point_col = None
        for c in df.columns:
            c_upper = str(c).upper()
            if any(term in c_upper for term in ['ACHIEVEMENT', 'POINT', 'STAGE', 'ACTIVITY']):
                point_col = c
                break
        if not point_col:
            if len(df.columns) > 1:
                point_col = df.columns[1]
            else:
                point_col = df.columns[0]
                
        # Filter matching rows
        matched_dict = {}
        for _, row in df.iterrows():
            raw_val = str(row[point_col]).strip()
            n_val = norm_key(raw_val)
            if n_val in norm_map:
                std_key = norm_map[n_val]
                row_copy = row.copy()
                row_copy[point_col] = std_key
                matched_dict[std_key] = row_copy
                
        if not matched_dict:
            return None
            
        # Build ordered dataframe
        ordered_rows = []
        for k in ['TCF1_VIN_GENERATION', 'TCF2_VIN_GENERATION', 'TCF1-DROP', 'TCF2-DROP']:
            if k in matched_dict:
                ordered_rows.append(matched_dict[k])
                
        res_df = pd.DataFrame(ordered_rows)
        
        # Remove SR.NO column if present to keep table clean
        cols_to_keep = [c for c in res_df.columns if not re.match(r'^(SR\.?\s*NO|S\.?\s*NO)$', str(c).strip(), re.I)]
        res_df = res_df[cols_to_keep].reset_index(drop=True)
        
        # Standardize point col name
        if point_col in res_df.columns and point_col != 'ACHIEVEMENT POINT':
            res_df.rename(columns={point_col: 'ACHIEVEMENT POINT'}, inplace=True)
            
        return res_df
    except Exception as e:
        print(f"Error loading hourly production: {e}")
        return None

