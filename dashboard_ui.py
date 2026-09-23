"""Navigation, report identities and on-demand downloads for the PPC dashboard."""
import datetime
import hashlib
import io
import os
import time

import pandas as pd
import streamlit as st

PAGES = ['Overview', 'Summary & Excel Reports', 'Cockpit & Wiring Shortages',
         'TCF1 Line', 'TCF2 Line', 'Total Float & Search', 'Quality Holds',
         'Control Panel', 'Telegram Dispatcher']
IST = datetime.timezone(datetime.timedelta(hours=5, minutes=30))


def invalidate_report():
    """A pending input change must not leave an old snapshot available to timers."""
    for key in ('_report_snapshot', '_summary_snapshot', '_prepared_exports'):
        st.session_state.pop(key, None)


def fingerprint(value):
    """Content identity; never use a buffer cursor or DB metadata write time."""
    digest = hashlib.sha256()
    def add(item):
        if isinstance(item, pd.DataFrame):
            digest.update(repr((list(item.columns), list(map(str, item.dtypes)))).encode())
            digest.update(pd.util.hash_pandas_object(item, index=True).values.tobytes())
        elif isinstance(item, dict):
            for key in sorted(item, key=str):
                add(key)
                add(item[key])
        elif isinstance(item, (tuple, list)):
            for element in item:
                add(element)
        elif hasattr(item, 'content_id'):
            add(item.content_id)
        elif hasattr(item, 'getvalue'):
            digest.update(item.getvalue())
        elif isinstance(item, (str, os.PathLike)) and os.path.isfile(item):
            stat = os.stat(item)
            digest.update(repr((os.path.abspath(item), stat.st_size, stat.st_mtime_ns)).encode())
        else:
            digest.update(repr(item).encode())
        digest.update(b'\x00')
    add(value)
    return digest.hexdigest()


def setup_shell():
    st.set_page_config(page_title='PPC Production Dashboard', page_icon='🚗',
                       layout='wide', initial_sidebar_state='expanded')
    # Preserve filter widget values when their page is not rendered.
    for key in list(st.session_state):
        if any(word in key for word in ('filter', 'search', 'view_mode', 'subview', 'detail_columns')):
            st.session_state[key] = st.session_state[key]
    st.session_state.setdefault('theme', '☀️ White Theme')
    with st.sidebar:
        st.title('PPC Dashboard')
        page = st.radio('Workspace', PAGES, key='dashboard_page', label_visibility='collapsed')
        st.divider()
        st.selectbox('Appearance', ['☀️ White Theme', '🌙 Dark Theme'], key='theme')
        if st.button('Refresh source files', use_container_width=True):
            st.session_state.pop('_file_registry', None)
            st.session_state.pop('last_onedrive_sync', None)
            st.session_state.pop('_report_snapshot', None)
            st.session_state['run_report'] = True
            st.rerun()
        st.caption('Stock and upload settings are in Control Panel.')
    dark = st.session_state.theme == '🌙 Dark Theme'
    bg, card, text, border = ('#0f172a','#1e293b','#f1f5f9','#334155') if dark else ('#f6f8fb','#ffffff','#172033','#e2e8f0')
    st.markdown(f'''<style>
    html, body, [class*="css"] {{font-family: system-ui, -apple-system, "Segoe UI", sans-serif;}}
    .stApp {{background:{bg}; color:{text};}}
    [data-testid="stSidebar"] {{background:{card};}}
    .block-container {{padding-top:1.3rem; padding-bottom:2rem;}}
    h1 {{font-size:1.65rem !important;}} h2 {{font-size:1.3rem !important;}}
    h3 {{font-size:1.08rem !important;}}
    [data-testid="stMetric"] {{background:{card};border:1px solid {border};border-radius:9px;padding:12px;}}
    [data-testid="stMetricValue"] {{font-size:1.55rem;}}
    [data-testid="stDataFrame"] {{border:1px solid {border};border-radius:8px;}}
    @media(max-width:700px) {{
      .block-container {{padding:1rem .7rem;}}
      [data-testid="stHorizontalBlock"] {{flex-wrap:wrap;gap:.6rem;}}
      [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {{min-width:45% !important;flex:1 1 45% !important;}}
    }}
    </style>''', unsafe_allow_html=True)
    st.title(page)
    return page


def cached_download(*, label, builder, file_name, key, version=None, **kwargs):
    """Builders run only on request. Invalidate when inputs or export filters change."""
    identity = fingerprint((st.session_state.get('_report_key'), version))
    cache = st.session_state.setdefault('_prepared_exports', {})
    entry = cache.get(key)
    if entry is not None and entry[0] != identity:
        cache.pop(key, None)
        entry = None
    if st.button('Prepare · ' + label.replace('📥 ', '').replace('⬇️ ', ''), key='prepare_' + key):
        with st.spinner('Preparing Excel file…'):
            data = builder()
        cache[key] = (identity, data)
        entry = cache[key]
        # Keep memory bounded; each report refresh also clears this cache.
        while len(cache) > 12:
            cache.pop(next(iter(cache)))
    if entry is not None:
        st.download_button(label, data=entry[1], file_name=file_name, key=key,
                           on_click='ignore', **kwargs)


def table_workbook(sheets):
    from openpyxl.styles import Alignment, Font, PatternFill
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, index=False, sheet_name=name)
            ws = writer.sheets[name]
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', fgColor='1D4ED8')
                cell.alignment = Alignment(wrap_text=True)
            for column in ws.columns:
                ws.column_dimensions[column[0].column_letter].width = min(48, max(14, max(len(str(c.value or '')) for c in column) + 2))
    return output.getvalue()


def file_registry(directory, detector):
    """Only scan on interval or explicit refresh; content parsers have their own cache."""
    cached = st.session_state.get('_file_registry')
    if not cached or cached[0] != directory or time.monotonic() - cached[1] > 60:
        sources = detector(directory)
        cached = (directory, time.monotonic(), sources, fingerprint(sources))
        st.session_state['_file_registry'] = cached
    return cached[2].copy()


def clear_stock_drafts():
    for key in ('engine_stock_editor', 'nova_stock_editor'):
        st.session_state.pop(key, None)


def show_data_health(sources):
    rows = []
    for category, source in sources.items():
        if isinstance(source, str) and os.path.isfile(source):
            timestamp = datetime.datetime.fromtimestamp(os.path.getmtime(source), IST).strftime('%d %b %H:%M IST')
            name = os.path.basename(source)
        else:
            name = getattr(source, 'name', str(source) if isinstance(source, str) else 'Workbook sheet')
            timestamp = st.session_state.get('upload_time_' + category, 'Not supplied')
        rows.append({'Report': category.replace('_', ' '), 'Source': name, 'File modified / received': timestamp})
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
    st.caption('File modified / received time is not necessarily the source report’s production cutoff time.')


def overview(namespace, sources):
    frames = [namespace['tcf1_alloc_df'], namespace['tcf2_alloc_df']]
    queue = pd.concat(frames, ignore_index=True)
    status = queue.get('STATUS', pd.Series(dtype=str))
    drops = sum(int(df['VIN_Count'].sum()) if 'VIN_Count' in df else len(df)
                for df in [namespace['tcf1_drops'], namespace['tcf2_drops']] if df is not None)
    holds = namespace['pbs_on_hold']
    metrics = [('VIN generated · both lines', drops), ('PBS ready', int(status.eq('✅ Ready for TCF').sum())),
               ('PBS material blocked', int(status.eq('🚫 Blocked').sum())), ('PBS quality holds', len(holds)),
               ('PBS BOM issues', int(status.str.startswith('⚠️', na=False).sum()))]
    for column, (label, value) in zip(st.columns(5), metrics):
        column.metric(label, value)
    st.subheader('Line status')
    rows = []
    for line, frame in zip(['TCF1', 'TCF2'], frames):
        values = frame.get('STATUS', pd.Series(dtype=str))
        rows.append({'Line': line, 'Ready': int(values.eq('✅ Ready for TCF').sum()),
                     'Blocked': int(values.eq('🚫 Blocked').sum()), 'BOM issues': int(values.str.startswith('⚠️',na=False).sum())})
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
    st.subheader('Main blocking reasons · PBS')
    blocked = queue[status.eq('🚫 Blocked')] if not queue.empty else queue
    if not blocked.empty:
        counts = blocked['BLOCKING_REASON'].fillna('Unknown').value_counts().head(8).rename_axis('Reason').reset_index(name='Affected cabs')
        st.dataframe(counts, hide_index=True, use_container_width=True)
    else:
        st.info('No PBS material blocking reasons in this report.')
    st.subheader('Find a vehicle')
    query = st.text_input('BIW / VIN / vehicle code', key='overview_search')
    data = namespace['temp_float_df']
    if query.strip() and not data.empty:
        mask = pd.Series(False,index=data.index)
        for col in ['BIW NUMBER','VIN','VEHICLE CODE']:
            if col in data: mask |= data[col].astype(str).str.contains(query.strip(),case=False,regex=False,na=False)
        columns = [c for c in ['BIW NUMBER','VIN','PRODUCT','SHOP','Stage','Status','Blocking Reason'] if c in data]
        st.dataframe(data.loc[mask,columns], hide_index=True,use_container_width=True)
    with st.expander('Source files and freshness'):
        show_data_health(sources)
