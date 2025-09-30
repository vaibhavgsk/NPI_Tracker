import pandas as pd
import plotly.graph_objects as go
import re
import math
from html import unescape
from pathlib import Path

df_JP = pd.read_excel("Template.xlsx", sheet_name='Japan')
df_CN = pd.read_excel("Template.xlsx", sheet_name='China')
df_CA = pd.read_excel("Template.xlsx", sheet_name='Canada')
df_AU = pd.read_excel("Template.xlsx", sheet_name='AUSTRALIA & Viiv')
df_KR = pd.read_excel("Template.xlsx", sheet_name='Korea')
df_TW = pd.read_excel("Template.xlsx", sheet_name='Taiwan')
df_HK = pd.read_excel("Template.xlsx", sheet_name='Hong Kong')
df_SG = pd.read_excel("Template.xlsx", sheet_name='Singapore')
df_NZ = pd.read_excel("Template.xlsx", sheet_name='Newzealand')

df_base = pd.concat(
    [df_JP, df_CN, df_CA, df_AU, df_KR, df_TW, df_HK, df_SG, df_NZ],
    ignore_index=True
)

df = df_base.copy()
df['Quarter'] = df.get('Quarter', pd.Series(['']*len(df))).astype(str).str.strip().replace({'nan':'', 'None':''}).str.upper()
df['Year'] = df.get('Year', pd.Series(['']*len(df))).astype(str).str.strip().replace({'nan':'', 'None':''})

def to_year_int(y):
    s = str(y).strip()
    return int(s) if re.match(r'^\d{4}$', s) else None

df['Year_int'] = df['Year'].apply(to_year_int)

def std_quarter(q):
    q = str(q).strip().upper()
    if re.match(r'^Q[1-4]$', q):
        return q
    m = re.search(r'(Q[1-4])', q)
    return m.group(1) if m else ''

df['Quarter_std'] = df['Quarter'].apply(std_quarter)

years = sorted({y for y in df['Year_int'].unique() if isinstance(y, int)})
if not years:
    yrs = set()
    for val in df['Quarter'].dropna().astype(str).unique():
        m = re.search(r"(\d{4})$", str(val))
        if m:
            yrs.add(int(m.group(1)))
    years = sorted(yrs) if yrs else [2025, 2026, 2027, 2028]

quarter_list = ['Q1', 'Q2', 'Q3', 'Q4']
index_tuples = [(y, q) for y in years for q in quarter_list]
idx = pd.MultiIndex.from_tuples(index_tuples, names=['Year_val', 'Quarter_val'])

def make_brand_info_html_and_plain(row):
    brand = str(row.get('Brand','') or '').strip()
    ra_best = str(row.get('RA Approval Month (Best)','') or '').strip()
    ra_base = str(row.get('RA Approval Month (Base)','') or '').strip()
    lm_best = str(row.get('Launch Month (Best)','') or '').strip()
    lm_base = str(row.get('Launch Month (Base)','') or '').strip()
    status = str(row.get('Status','') or '').strip()

    lines_html = []
    lines_plain = []
    lines_html.append(f"<b>Brand:</b> {brand}" if brand else "<b>Brand:</b>")
    lines_plain.append(f"Brand: {brand}" if brand else "Brand:")
    lines_html.append(f"<b>RA Approval (Best):</b> {ra_best}" if ra_best else "<b>RA Approval (Best):</b>")
    lines_plain.append(f"RA Approval (Best): {ra_best}" if ra_best else "RA Approval (Best):")
    lines_html.append(f"<b>RA Approval (Base):</b> {ra_base}" if ra_base else "<b>RA Approval (Base):</b>")
    lines_plain.append(f"RA Approval (Base): {ra_base}" if ra_base else "RA Approval (Base):")
    lines_html.append(f"<b>Launch (Best):</b> {lm_best}" if lm_best else "<b>Launch (Best):</b>")
    lines_plain.append(f"Launch (Best): {lm_best}" if lm_best else "Launch (Best):")
    lines_html.append(f"<b>Launch (Base):</b> {lm_base}" if lm_base else "<b>Launch (Base):</b>")
    lines_plain.append(f"Launch (Base): {lm_base}" if lm_base else "Launch (Base):")
    lines_html.append(f"<b>Status:</b> {status}" if status else "<b>Status:</b>")
    lines_plain.append(f"Status: {status}" if status else "Status:")

    html = "<br>".join(lines_html)
    plain = " ; ".join([p for p in lines_plain if p])
    return pd.Series({'Brand_Info_html': html, 'Brand_Info_plain': plain})

bi = df.apply(make_brand_info_html_and_plain, axis=1)
df = pd.concat([df, bi], axis=1)
df['brand_col'] = df.get('LOC', pd.Series(['']*len(df))).astype(str).str.strip()

agg_df = df[df['Year_int'].notna() & df['Quarter_std'].ne('')].copy()
grouped = agg_df.groupby(['Year_int', 'Quarter_std', 'brand_col']).agg({
    'Brand_Info_html': lambda s: "<br><br>".join(s.dropna().astype(str)),
    'Brand_Info_plain': lambda s: " ; ".join(s.dropna().astype(str))
}).reset_index()

loc_cols = []
for loc in df['brand_col'].tolist():
    if loc not in loc_cols:
        loc_cols.append(loc)

wide_cols = ['Year', 'Quarter'] + loc_cols
wide = pd.DataFrame('', index=idx, columns=wide_cols)

for (y, q) in index_tuples:
    wide.at[(y, q), 'Year'] = str(y) if q == 'Q2' else ''
    wide.at[(y, q), 'Quarter'] = q

plain_map = {}
for _, r in grouped.iterrows():
    y = int(r['Year_int'])
    q = r['Quarter_std']
    loc = r['brand_col']
    html_val = r['Brand_Info_html'] or ''
    plain_val = r['Brand_Info_plain'] or ''
    if (y, q) not in wide.index:
        continue
    if loc not in wide.columns:
        wide[loc] = ''
    prev_html = wide.at[(y, q), loc]
    wide.at[(y, q), loc] = (prev_html + "<br><br>" + html_val) if prev_html else html_val
    prev_plain = plain_map.get((y, q, loc), '')
    plain_map[(y, q, loc)] = (prev_plain + " ; " + plain_val) if prev_plain else plain_val

def status_color_for_plain_text(text):
    t = str(text).upper()
    if 'LAUNCHED' in t or re.search(r'\bLAUNCH\b', t):
        return '#c8f7d7'
    if 'IN PROGRESS' in t or 'INPROGRESS' in t or re.search(r'\bIN\s*PROG', t):
        return '#d9d9d9'
    if 'TBC' in t or 'TBD' in t:
        return '#e6f0ff'
    if 'APPROVAL' in t or 'APPROVED' in t:
        return '#fff3bf'
    return 'white'

fill_color_per_col = []
for col_idx, col in enumerate(wide.columns):
    col_colors = []
    for i, row_val in enumerate(wide[col].tolist()):
        if col in ['Year', 'Quarter']:
            col_colors.append('#f0f0f0')
        else:
            (y_val, q_val) = wide.index[i]
            plain_for_cell = plain_map.get((y_val, q_val, col), '')
            if not plain_for_cell:
                plain_for_cell = re.sub('<[^<]+?>', '', str(row_val or ''))
            col_colors.append(status_color_for_plain_text(plain_for_cell))
    fill_color_per_col.append(col_colors)

def html_to_visible_lines(html_text):
    if html_text is None:
        return ['']
    s = str(html_text).replace('<br><br>', '\n\n').replace('<br/>', '\n').replace('<br>', '\n')
    s_no_tags = re.sub('<[^<]+?>', '', s)
    s_no_entities = unescape(s_no_tags)
    lines = [ln.rstrip() for ln in s_no_entities.splitlines()]
    if not any(lines):
        return ['']
    return lines

max_chars_per_col = []
for col in wide.columns:
    max_chars = 0
    for val in wide[col].tolist():
        lines = html_to_visible_lines(val)
        for ln in lines:
            ln_len = len(ln)
            if ln_len > max_chars:
                max_chars = ln_len
    max_chars_per_col.append(max_chars)

char_to_px = 7
col_widths = []
for col_idx, col in enumerate(wide.columns):
    if col in ['Year']:
        px = 80
    elif col in ['Quarter']:
        px = 100
    else:
        max_chars = max_chars_per_col[col_idx]
        px = max(180, min(900, int(math.ceil(max_chars * char_to_px))))
    col_widths.append(px)

chars_per_line_per_col = [max(1, int(w // char_to_px)) for w in col_widths]

rows_line_counts = []
max_lines_needed = 1
for row_idx in range(len(wide.index)):
    row_lines_needed = 1
    for col_idx, col in enumerate(wide.columns):
        raw_val = wide.iloc[row_idx, col_idx]
        visible_lines = html_to_visible_lines(raw_val)
        chars_per_line = chars_per_line_per_col[col_idx]
        total_lines = 0
        for vl in visible_lines:
            if not vl:
                total_lines += 1
            else:
                total_lines += math.ceil(len(vl) / chars_per_line)
        if total_lines > row_lines_needed:
            row_lines_needed = total_lines
    rows_line_counts.append(row_lines_needed)
    if row_lines_needed > max_lines_needed:
        max_lines_needed = row_lines_needed

px_per_line = 16
row_height = max(38, px_per_line * max_lines_needed + 8)
table_total_width = sum(col_widths) + 40

left_cols = ['Year', 'Quarter']
right_cols = [c for c in wide.columns if c not in left_cols]

left_values = [[''] * len(wide.index) for _ in left_cols]

right_values = []
for c in right_cols:
    col_list = []
    for val in wide[c].tolist():
        col_list.append(val if val is not None else '')
    right_values.append(col_list)

left_fill = fill_color_per_col[0:len(left_cols)]
right_fill = fill_color_per_col[len(left_cols):]
left_widths = col_widths[0:len(left_cols)]
right_widths = col_widths[len(left_cols):]

if not right_cols:
    right_cols = ['_empty_']
    right_values = [[''] * len(wide.index)]
    right_fill = [['white'] * len(wide.index)]
    right_widths = [200]

total_px = sum(col_widths) if sum(col_widths) > 0 else 1
left_domain_width = sum(left_widths) / total_px

table_left = go.Table(
    domain=dict(x=[0.0, left_domain_width], y=[0, 1]),
    header=dict(
        values=[f"<b>{c}</b>" for c in left_cols],
        fill_color='#d95f02',
        align='center',
        font=dict(color='white', size=16),
        height=56
    ),
    cells=dict(
        values=left_values,
        fill_color=left_fill,
        align=['center'] * len(left_cols),
        height=row_height,
        font=dict(color='black', size=18),
    ),
    columnwidth=left_widths
)

table_right = go.Table(
    domain=dict(x=[left_domain_width, 1.0], y=[0, 1]),
    header=dict(
        values=[f"<b>{c}</b>" for c in right_cols],
        fill_color='#d95f02',
        align='center',
        font=dict(color='white', size=14),
        height=56
    ),
    cells=dict(
        values=right_values,
        fill_color=right_fill,
        align=['left'] * len(right_cols),
        height=row_height,
        font=dict(color='black', size=11),
    ),
    columnwidth=right_widths
)

fig_height = 200 + (row_height * len(wide.index))
fig = go.Figure(data=[table_left, table_right])
fig.update_layout(
    width=table_total_width,
    height=fig_height,
    margin=dict(l=10, r=10, t=10, b=10)
)

annotations = []
header_px = 56

cum = 0
left_centers = []
for w in left_widths:
    center_px = cum + (w / 2.0)
    left_centers.append(center_px / table_total_width)
    cum += w

n_rows = len(wide.index)
for row_idx in range(n_rows):
    data_top = 1.0 - (header_px / fig_height)
    row_center_offset = (row_idx + 0.5) * (row_height / fig_height)
    y_center = data_top - row_center_offset

    year_txt = wide.iloc[row_idx, 0]
    if year_txt and str(year_txt).strip():
        annotations.append(dict(
            x=left_centers[0],
            y=y_center,
            xref='paper', yref='paper',
            text=f"<b>{str(year_txt).strip()}</b>",
            showarrow=False,
            xanchor='center',
            yanchor='middle',
            font=dict(size=16, color='black')
        ))

    quarter_txt = wide.iloc[row_idx, 1]
    if quarter_txt and str(quarter_txt).strip():
        annotations.append(dict(
            x=left_centers[1],
            y=y_center,
            xref='paper', yref='paper',
            text=f"<b>{str(quarter_txt).strip()}</b>",
            showarrow=False,
            xanchor='center',
            yanchor='middle',
            font=dict(size=14, color='black')
        ))

fig.update_layout(annotations=annotations)
fig.show()

out = Path('NPI_Tracker.html')
fig.write_html(out)
print("Saved table to:", out)
