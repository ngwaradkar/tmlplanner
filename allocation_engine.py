import pandas as pd
import numpy as np

def _bom_lookup(bom):
    if bom is None or bom.empty:
        return {}
    return {row['Short Vehicle Code']: row for row in
            bom.drop_duplicates('Short Vehicle Code', keep='first').to_dict('records')}

def calculate_true_stock(shift_start_stock, tcf_drops, bom, bom_part_col):
    """
    Computes True Current Stock = Shift Start Stock - Consumed Parts.
    - shift_start_stock: dict of {part_number: qty}
    - tcf_drops: DataFrame of cabs built this shift
    - bom: DataFrame of BOM mappings (Short VC -> Engine, Cockpit, Front Wiring)
    - bom_part_col: column in BOM corresponding to this part type ('Engine', 'Cockpit', 'Front Wiring')
    
    Returns:
      - true_stock: dict of {part_number: true_qty}
      - consumed: dict of {part_number: consumed_qty}
      - warnings: list of strings (e.g. part number went negative)
    """
    if shift_start_stock is None:
        return None, {}, ['Stock file not loaded.']
    lookup = _bom_lookup(bom)
    true_stock = shift_start_stock.copy()
    consumed = {part: 0 for part in shift_start_stock}
    warnings = []
    
    # If no tcf_drops or bom, true_stock is just shift_start_stock
    if tcf_drops is None or tcf_drops.empty or bom is None or bom.empty:
        return true_stock, consumed, warnings
        
    for idx, row in tcf_drops.iterrows():
        full_vc = row.get('VEHICLE CODE') if pd.notna(row.get('VEHICLE CODE')) else row.get('VC')
        if pd.isna(full_vc) or not full_vc:
            continue
        short_vc = str(full_vc).strip()[:9]
        
        # Look up in BOM
        bom_entry = lookup.get(short_vc)
        if bom_entry is None:
            continue
            
        part_no = bom_entry.get(bom_part_col)
        if not part_no or str(part_no).strip() in ['0', 'None', 'nan']:
            continue
            
        cnt = int(row.get('VIN_Count', 1)) if pd.notna(row.get('VIN_Count')) and str(row.get('VIN_Count')).isdigit() else 1
        if part_no in true_stock:
            true_stock[part_no] -= cnt
            consumed[part_no] += cnt
        else:
            consumed[part_no] = consumed.get(part_no, 0) + cnt
            
    # Check for negative true stock
    for part, qty in true_stock.items():
        if qty < 0:
            warnings.append(f"Part {part} has negative true current stock ({qty}) due to backflushing.")
            
    return true_stock, consumed, warnings

def _is_model_trim_matched(cab_model, cab_sales_desc, target_model, target_trims):
    t_mod = str(target_model).strip().upper()
    c_mod = str(cab_model).strip().upper() if cab_model else ''
    c_desc = str(cab_sales_desc).strip().upper() if cab_sales_desc else ''
    
    model_match = False
    if t_mod == 'HARRIER / SAFARI':
        if any(k in c_mod or k in c_desc for k in ['HARRIER', 'SAFARI', 'GRAVITAS', 'Q5']):
            model_match = True
    elif t_mod == 'HARRIER.EV':
        if any(k in c_mod or k in c_desc for k in ['HARRIER.EV', 'ETURNA']):
            model_match = True
    elif t_mod == 'PUNCH.EV':
        if any(k in c_mod or k in c_desc for k in ['PUNCH.EV', 'NOVA']):
            model_match = True
    elif t_mod == 'PUNCH':
        if any(k in c_mod or k in c_desc for k in ['PUNCH', 'HORNBILL']) and 'EV' not in c_mod and 'NOVA' not in c_mod and 'PUNCH.EV' not in c_desc:
            model_match = True
    else:
        if t_mod in c_mod or t_mod in c_desc:
            model_match = True

    if not model_match:
        return False

    t_trims = str(target_trims).strip().upper()
    if not t_trims or 'ALL TRIMS' in t_trims:
        return True

    trims_list = [t.strip() for t in t_trims.split(',') if t.strip()]
    for t_val in trims_list:
        if t_val in c_desc:
            return True

    return False


def run_allocation(pbs_queue, bom, true_engine, true_cockpit, true_wiring, true_nova=None, model_shortages=None):
    """
    Runs the FIFO allocation loop for PBS cabs.
    - pbs_queue: DataFrame of cabs in PBS (sorted FIFO by PBS LIFT, HOLD BY is null)
    - bom: BOM DataFrame
    - true_engine: dict of true stock for Engines (None if not available)
    - true_cockpit: dict of true stock for Cockpits (None if not available)
    - true_wiring: dict of true stock for Wiring (None if not available)
    - true_nova: dict of Punch EV material stocks (Battery, Combo, Tube Frame(Craddle), Subframe, RTB)
    - model_shortages: list of dicts [{'Model': '...', 'Trims': '...', 'Part Name': '...', 'Stock': int}]
    
    Returns:
      - results: list of dicts representing allocated cabs
      - final_stocks: dict containing final virtual stocks
    """
    lookup = _bom_lookup(bom)
    # Create working copies of stock pools
    virt_engine = true_engine.copy() if true_engine is not None else None
    virt_cockpit = true_cockpit.copy() if true_cockpit is not None else None
    virt_wiring = true_wiring.copy() if true_wiring is not None else None
    virt_nova = true_nova.copy() if true_nova is not None else None
    virt_model_shortages = [dict(item) for item in model_shortages] if model_shortages else []
    
    results = []
    
    # If no cabs in queue, return empty
    if pbs_queue is None or pbs_queue.empty:
        return [], {
            'engine': virt_engine,
            'cockpit': virt_cockpit,
            'wiring': virt_wiring,
            'nova': virt_nova,
            'model_shortages': virt_model_shortages
        }
        
    # Process cabs in FIFO order
    for idx, row in pbs_queue.iterrows():
        biw_num = row.get('BIW NUMBER')
        vin = row.get('VIN')
        full_vc = row.get('VEHICLE CODE') if pd.notna(row.get('VEHICLE CODE')) else row.get('VC')
        pbs_lift = row.get('PBS LIFT')
        colour = row.get('COLOUR')
        product = row.get('PRODUCT')
        sales_desc = row.get('SALES DESCRIPTION')
        shop = row.get('SHOP')
        
        short_vc = str(full_vc).strip()[:9]
        
        # BOM Lookup
        bom_entry = lookup.get(short_vc)
        
        if bom_entry is None:
            results.append({
                'BIW NUMBER': biw_num,
                'VIN': vin,
                'VEHICLE CODE': full_vc,
                'Short VC': short_vc,
                'PBS LIFT': pbs_lift,
                'COLOUR': colour,
                'PRODUCT': product,
                'SALES DESCRIPTION': sales_desc,
                'SHOP': shop,
                'STATUS': '⚠️ Unknown VC',
                'BLOCKING_REASON': f"VC prefix '{short_vc}' not found in BOM.",
                'Engine_Part': None,
                'Cockpit_Part': None,
                'Wiring_Part': None,
                'Engine_Stock_After': None,
                'Cockpit_Stock_After': None,
                'Wiring_Stock_After': None
            })
            continue
            
        eng_part = bom_entry.get('Engine')
        ck_part = bom_entry.get('Cockpit')
        wh_part = bom_entry.get('Front Wiring')
        
        eng_part_str = str(eng_part).strip() if eng_part else None
        ck_part_str = str(ck_part).strip() if ck_part else None
        wh_part_str = str(wh_part).strip() if wh_part else None
        
        # Check for incomplete BOM (any required part is None or '0')
        is_bat_bom = (product and ('EV' in str(product).upper() or 'NOVA' in str(product).upper()))
        incomplete = []
        if not eng_part_str or eng_part_str == '0':
            incomplete.append('Battery' if is_bat_bom else 'Engine')
        if not ck_part_str or ck_part_str == '0':
            incomplete.append('Cockpit')
        if not wh_part_str or wh_part_str == '0':
            incomplete.append('Front Wiring')
            
        if incomplete:
            results.append({
                'BIW NUMBER': biw_num,
                'VIN': vin,
                'VEHICLE CODE': full_vc,
                'Short VC': short_vc,
                'PBS LIFT': pbs_lift,
                'COLOUR': colour,
                'PRODUCT': product,
                'SALES DESCRIPTION': sales_desc,
                'SHOP': shop,
                'STATUS': '⚠️ BOM Incomplete',
                'BLOCKING_REASON': f"Incomplete BOM for parts: {', '.join(incomplete)}",
                'Engine_Part': eng_part_str,
                'Cockpit_Part': ck_part_str,
                'Wiring_Part': wh_part_str,
                'Engine_Stock_After': None,
                'Cockpit_Stock_After': None,
                'Wiring_Stock_After': None
            })
            continue
            
        # Check stock availability
        has_engine = True
        has_cockpit = True
        has_wiring = True
        has_model_shortage = False
        
        shortages = []
        
        # Engine / Battery / Punch EV stock check
        if eng_part_str == '546816111212' and virt_nova:
            nova_shortages = []
            for mat_name, mat_qty in virt_nova.items():
                if mat_qty <= 0:
                    nova_shortages.append(f"{mat_name} (Stock: {mat_qty})")
            if nova_shortages:
                has_engine = False
                shortages.extend(nova_shortages)
        elif virt_engine is not None:
            stock = virt_engine.get(eng_part_str, 0)
            if stock <= 0:
                has_engine = False
                is_battery = eng_part_str in ['546816111212', '547380400103'] or (product and ('EV' in str(product).upper() or 'NOVA' in str(product).upper()))
                part_lbl = 'Battery' if is_battery else 'Engine'
                shortages.append(f"{part_lbl} {eng_part_str} (Stock: {stock})")
                
        # Cockpit stock check
        if virt_cockpit is not None and ck_part_str in virt_cockpit:
            stock = virt_cockpit.get(ck_part_str, 0)
            if stock <= 0:
                has_cockpit = False
                shortages.append(f"Cockpit {ck_part_str} (Stock: {stock})")
                
        # Wiring stock check
        if virt_wiring is not None and wh_part_str in virt_wiring:
            stock = virt_wiring.get(wh_part_str, 0)
            if stock <= 0:
                has_wiring = False
                shortages.append(f"Wiring {wh_part_str} (Stock: {stock})")
                
        # Model-Wise Shortages stock check
        matched_ms_items = []
        if virt_model_shortages:
            for ms_item in virt_model_shortages:
                if _is_model_trim_matched(product, sales_desc, ms_item['Model'], ms_item.get('Trims', 'All Trims')):
                    matched_ms_items.append(ms_item)
                    if ms_item['Stock'] <= 0:
                        has_model_shortage = True
                        shortages.append(f"{ms_item['Part Name']} (Stock: {ms_item['Stock']})")

        # If all available, allocate
        if has_engine and has_cockpit and has_wiring and not has_model_shortage:
            # Decrement stocks
            if virt_engine is not None and eng_part_str in virt_engine:
                virt_engine[eng_part_str] -= 1
            if eng_part_str == '546816111212' and virt_nova:
                for mat_name in virt_nova:
                    virt_nova[mat_name] -= 1
            if virt_cockpit is not None and ck_part_str in virt_cockpit:
                virt_cockpit[ck_part_str] -= 1
            if virt_wiring is not None and wh_part_str in virt_wiring:
                virt_wiring[wh_part_str] -= 1
            for ms_item in matched_ms_items:
                ms_item['Stock'] -= 1
                
            results.append({
                'BIW NUMBER': biw_num,
                'VIN': vin,
                'VEHICLE CODE': full_vc,
                'Short VC': short_vc,
                'PBS LIFT': pbs_lift,
                'COLOUR': colour,
                'PRODUCT': product,
                'SALES DESCRIPTION': sales_desc,
                'SHOP': shop,
                'STATUS': '✅ Ready for TCF',
                'BLOCKING_REASON': None,
                'Engine_Part': eng_part_str,
                'Cockpit_Part': ck_part_str,
                'Wiring_Part': wh_part_str,
                'Engine_Stock_After': virt_engine.get(eng_part_str) if virt_engine is not None else None,
                'Cockpit_Stock_After': virt_cockpit.get(ck_part_str) if virt_cockpit is not None else None,
                'Wiring_Stock_After': virt_wiring.get(wh_part_str) if virt_wiring is not None else None
            })
        else:
            results.append({
                'BIW NUMBER': biw_num,
                'VIN': vin,
                'VEHICLE CODE': full_vc,
                'Short VC': short_vc,
                'PBS LIFT': pbs_lift,
                'COLOUR': colour,
                'PRODUCT': product,
                'SALES DESCRIPTION': sales_desc,
                'SHOP': shop,
                'STATUS': '🚫 Blocked',
                'BLOCKING_REASON': f"Shortage: {', '.join(shortages)}",
                'Engine_Part': eng_part_str,
                'Cockpit_Part': ck_part_str,
                'Wiring_Part': wh_part_str,
                'Engine_Stock_After': virt_engine.get(eng_part_str) if virt_engine is not None else None,
                'Cockpit_Stock_After': virt_cockpit.get(ck_part_str) if virt_cockpit is not None else None,
                'Wiring_Stock_After': virt_wiring.get(wh_part_str) if virt_wiring is not None else None
            })
            
    return results, {
        'engine': virt_engine,
        'cockpit': virt_cockpit,
        'wiring': virt_wiring
    }

def get_paint_float_stages(df_float):
    """
    Classifies each cab in the float report into its current stage in the paint flow.
    Flow order: BIW LIFTING -> PTCED -> SEALANT -> TOPCOAT -> PBS LIFT (closest to TCF)
    """
    stages = []
    
    for idx, row in df_float.iterrows():
        biw_lift = row.get('BIW LIFTING')
        ptced = row.get('PTCED')
        sealant = row.get('SEALANT')
        topcoat = row.get('TOPCOAT')
        pbs_lift = row.get('PBS LIFT')
        
        # Ordered classification (check closest to PBS first)
        if pd.notna(pbs_lift):
            stage = '1. PBS LIFT'
        elif pd.notna(topcoat):
            stage = '2. TOPCOAT'
        elif pd.notna(sealant):
            stage = '3. SEALANT'
        elif pd.notna(ptced):
            stage = '4. PTCED'
        elif pd.notna(biw_lift):
            stage = '5. BIW LIFTING'
        else:
            stage = '6. UNKNOWN'
            
        stages.append(stage)
        
    df_with_stage = df_float.copy()
    df_with_stage['Paint_Stage'] = stages
    return df_with_stage

def get_detailed_paint_summary_stage(row):
    """
    Classifies a float report cab into one of the 9 official Paint Shop Float Summary stages:
    1. PBS FLOAT
    2. PBS TO POLISHING
    3. POLISHING TO TOPCOAT
    4. TOPCOAT TO WETSANDING G ROOFBLACK
    5. TOPCOAT TO WETSANDING G FRESH
    6. WETSANDING G TO SEALANT
    7. PT ENTRY TO SEALANT
    8. BIW LIFTING G TO PT
    9. PT BYPASS
    """
    pbs_t = pd.to_datetime(row.get('PBS LIFT'), dayfirst=True, errors='coerce') if pd.notna(row.get('PBS LIFT')) else None
    topcoat_t = pd.to_datetime(row.get('TOPCOAT'), dayfirst=True, errors='coerce') if pd.notna(row.get('TOPCOAT')) else None
    sealant_t = pd.to_datetime(row.get('SEALANT'), dayfirst=True, errors='coerce') if pd.notna(row.get('SEALANT')) else None
    ptced_t = pd.to_datetime(row.get('PTCED'), dayfirst=True, errors='coerce') if pd.notna(row.get('PTCED')) else None
    biw_t = pd.to_datetime(row.get('BIW LIFTING'), dayfirst=True, errors='coerce') if pd.notna(row.get('BIW LIFTING')) else None
    
    colour = str(row.get('COLOUR', '')).strip().upper()
    is_dual = '-' in colour or '/' in colour or 'GBK' in colour or 'WHT' in colour
    
    if pd.notna(pbs_t):
        return 'PBS FLOAT'
    elif pd.notna(topcoat_t):
        t_hour = topcoat_t.hour + topcoat_t.minute / 60.0
        if t_hour <= 5.25 or (topcoat_t.date() < pd.Timestamp.now().date() and t_hour <= 12):
            return 'PBS TO POLISHING'
        else:
            return 'POLISHING TO TOPCOAT'
    elif pd.notna(sealant_t):
        s_hour = sealant_t.hour + sealant_t.minute / 60.0
        if s_hour <= 4.37 or (sealant_t.date() < pd.Timestamp.now().date() and s_hour <= 18):
            if is_dual and (str(row.get('BIW NUMBER','')) == '7012922' or 'ROOF' in colour):
                return 'TOPCOAT TO WETSANDING G ROOFBLACK'
            else:
                return 'TOPCOAT TO WETSANDING G FRESH'
        else:
            return 'WETSANDING G TO SEALANT'
    elif pd.notna(ptced_t):
        return 'PT ENTRY TO SEALANT'
    elif pd.notna(biw_t):
        return 'BIW LIFTING G TO PT'
    else:
        return 'PT BYPASS'

def calculate_stagewise_shortage(df_float_stages, bom, true_stocks):
    """
    Computes material requirements and shortages for each stage of the paint float.
    - df_float_stages: Float report with 'Paint_Stage' column and 'SHOP' column
    - bom: Master BOM DataFrame
    - true_stocks: dict containing {'engine': dict, 'cockpit': dict, 'wiring': dict} True Stock pools
    
    Returns:
      - shortage_report: DataFrame with columns: Stage, TCF Line, Aggregate Type, Part Number, Demand in Stage, Cumulative Demand, True Stock, Net Balance, Status
    """
    lookup = _bom_lookup(bom)
    # Sort order of stages (from closest to TCF to furthest)
    stage_order = ['1. PBS LIFT', '2. TOPCOAT', '3. SEALANT', '4. PTCED', '5. BIW LIFTING']
    
    # Accumulate demand per (stage, agg_type, part_number, shop)
    demand_counts = {}
    
    for idx, row in df_float_stages.iterrows():
        stage = row['Paint_Stage']
        if stage not in stage_order:
            continue
            
        shop = row.get('SHOP')
        if pd.isna(shop) or not str(shop).strip():
            shop = 'Unknown'
        else:
            shop = str(shop).strip()
            
        full_vc = row.get('VEHICLE CODE') if pd.notna(row.get('VEHICLE CODE')) else row.get('VC')
        short_vc = str(full_vc).strip()[:9] if pd.notna(full_vc) else ''
        
        # Look up in BOM
        bom_entry = lookup.get(short_vc)
        if bom_entry is None:
            continue
            
        
        parts = {
            'Engine': str(bom_entry.get('Engine')).strip() if bom_entry.get('Engine') else None,
            'Cockpit': str(bom_entry.get('Cockpit')).strip() if bom_entry.get('Cockpit') else None,
            'Front Wiring': str(bom_entry.get('Front Wiring')).strip() if bom_entry.get('Front Wiring') else None
        }
        
        for agg_type, part_no in parts.items():
            if not part_no or part_no in ['0', 'None', 'nan']:
                continue
                
            key = (stage, agg_type, part_no, shop)
            demand_counts[key] = demand_counts.get(key, 0) + 1
            
    # Calculate cumulative demand per (agg_type, part_no, shop)
    part_keys = set((agg_type, part_no, shop) for (_, agg_type, part_no, shop) in demand_counts.keys())
    
    report_rows = []
    
    # For each part number on each TCF Line, calculate demands across all stages
    for agg_type, part_no, shop in sorted(part_keys, key=lambda x: (x[2], x[0], x[1])):
        # Find stock
        stock = 0
        stock_loaded = False
        if agg_type == 'Engine' and true_stocks.get('engine') is not None:
            stock = true_stocks['engine'].get(part_no, 0)
            stock_loaded = True
        elif agg_type == 'Cockpit' and true_stocks.get('cockpit') is not None:
            stock = true_stocks['cockpit'].get(part_no, 0)
            stock_loaded = True
        elif agg_type == 'Front Wiring' and true_stocks.get('wiring') is not None:
            stock = true_stocks['wiring'].get(part_no, 0)
            stock_loaded = True
            
        cum_demand = 0
        for stage in stage_order:
            stage_demand = demand_counts.get((stage, agg_type, part_no, shop), 0)
            cum_demand += stage_demand
            
            balance = stock - cum_demand
            
            if stage_demand > 0 or cum_demand > 0:
                # Determine status
                if not stock_loaded:
                    status = "⚠️ Stock Not Loaded"
                elif balance < 0:
                    status = f"🚫 Shortage ({abs(balance)} units)"
                elif balance < 5:
                    status = "🟠 Low Stock Warning"
                else:
                    status = "🟢 Healthy"
                    
                display_agg = 'Battery' if (agg_type == 'Engine' and (part_no in ['546816111212', '547380400103'])) else agg_type
                report_rows.append({
                    'Stage': stage,
                    'TCF Line': shop,
                    'Aggregate Type': display_agg,
                    'Part Number': part_no,
                    'Stage Demand': stage_demand,
                    'Cumulative Demand': cum_demand,
                    'True Current Stock': stock if stock_loaded else None,
                    'Net Balance': balance if stock_loaded else None,
                    'Status': status
                })
                
    return pd.DataFrame(report_rows)


def find_missing_bom_vcs(df_float, bom):
    """
    Scans every cab currently in the float report and flags Short Vehicle Codes
    that either have no BOM row at all, or have a BOM row with one or more part
    numbers blank/'0'. This is the check the homepage alert/BOM-entry form is
    built on -- a cab whose BOM can't be resolved should never silently count
    as "Ready for TCF"; it should show up here instead.

    - df_float: Float report DataFrame with a 'VEHICLE CODE' (or 'VC') column
    - bom: Master BOM DataFrame (may be None if no BOM has been loaded yet)

    Returns a DataFrame with columns: Short VC, Status, Missing Parts, Cab Count,
    Sample Product, Sample VIN -- one row per distinct Short VC needing attention.
    Empty DataFrame means every cab in the float report has a complete BOM.
    """
    if df_float is None or df_float.empty:
        return pd.DataFrame()
    if bom is None:
        # No BOM loaded at all is a separate, already-surfaced condition
        # (Control Panel shows BOM as missing); don't duplicate that here.
        return pd.DataFrame()

    bom_lookup = {}
    for _, r in bom.iterrows():
        bom_lookup[str(r.get('Short Vehicle Code')).strip()] = r

    def _blank(v):
        if v is None:
            return True
        s = str(v).strip()
        return s == '' or s in ('0', 'None', 'nan')

    found = {}
    for _, row in df_float.iterrows():
        full_vc = row.get('VEHICLE CODE') if pd.notna(row.get('VEHICLE CODE')) else row.get('VC')
        if pd.isna(full_vc) or not str(full_vc).strip():
            continue
        short_vc = str(full_vc).strip()[:9]

        bom_entry = bom_lookup.get(short_vc)
        if bom_entry is None:
            status = '❌ Missing'
            missing_parts = ['Engine/Battery', 'Cockpit', 'Front Wiring']
        else:
            missing_parts = []
            if _blank(bom_entry.get('Engine')):
                missing_parts.append('Engine/Battery')
            if _blank(bom_entry.get('Cockpit')):
                missing_parts.append('Cockpit')
            if _blank(bom_entry.get('Front Wiring')):
                missing_parts.append('Front Wiring')
            if not missing_parts:
                continue
            status = '⚠️ Incomplete'

        if short_vc not in found:
            found[short_vc] = {
                'Short VC': short_vc,
                'Status': status,
                'Missing Parts': ', '.join(missing_parts),
                'Cab Count': 0,
                'Sample Product': row.get('PRODUCT') if pd.notna(row.get('PRODUCT')) else '',
                'Sample VIN': row.get('VIN') if pd.notna(row.get('VIN')) else ''
            }
        found[short_vc]['Cab Count'] += 1

    if not found:
        return pd.DataFrame()
    return pd.DataFrame(list(found.values())).sort_values('Cab Count', ascending=False).reset_index(drop=True)

