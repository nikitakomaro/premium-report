import streamlit as st
import pandas as pd
import io
import os
import warnings
import tempfile
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from bidi.algorithm import get_display

warnings.filterwarnings('ignore')

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="דוח פרמיה | Surense",
    page_icon="📊",
    layout="centered",
)

# ── RTL CSS ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    body, .stApp { direction: rtl; }
    .stFileUploader, .stDownloadButton, .stAlert, .stMetric,
    .stMarkdown, h1, h2, h3, p, label { text-align: right; direction: rtl; }
    .stDownloadButton button { width: 100%; font-size: 15px; padding: 10px; }
    .metric-box {
        background: #f0f4ff;
        border-radius: 10px;
        padding: 16px 20px;
        text-align: center;
        border: 1px solid #d0d8f0;
    }
    .metric-box .val { font-size: 2rem; font-weight: bold; color: #1F4E79; }
    .metric-box .lbl { font-size: 0.85rem; color: #555; margin-top: 4px; }
    .red   { color: #CC0000 !important; }
    .green { color: #1A5C1A !important; }
    .section-divider { border-top: 2px solid #e0e0e0; margin: 24px 0; }
</style>
""", unsafe_allow_html=True)

# ── Hebrew PDF helper ─────────────────────────────────────────────────────────
_font_registered = False
BASE_FONT = 'Helvetica'

def _register_font():
    global _font_registered, BASE_FONT
    if _font_registered:
        return
    for fp in ['/System/Library/Fonts/Supplemental/Arial Unicode.ttf',
                '/Library/Fonts/Arial Unicode.ttf',
                '/System/Library/Fonts/Supplemental/Arial.ttf']:
        if os.path.exists(fp):
            try:
                pdfmetrics.registerFont(TTFont('HebFont', fp))
                BASE_FONT = 'HebFont'
                break
            except:
                pass
    _font_registered = True

def rh(text):
    return get_display(str(text)) if text is not None else ''

# ── Core analysis ─────────────────────────────────────────────────────────────
COL_POLICY  = "מס' חשבון/פוליסה"
COL_ID      = 'מספר ת.ז'
COL_FNAME   = 'שם פרטי לקוח'
COL_LNAME   = 'שם משפחה לקוח'
COL_PREMIUM = 'סה"כ פרמיה'
COL_MFG     = 'יצרן'
COL_PRODUCT = 'מוצר'
COL_AGENT   = 'מת"ל'
SHEET       = 'מוצרי ביטוח'
THRESHOLD   = 15.0

thin = Side(style='thin', color='CCCCCC')
BORD = Border(left=thin, right=thin, top=thin, bottom=thin)
CTR  = Alignment(horizontal='center', vertical='center')
RGT  = Alignment(horizontal='right',  vertical='center')
DFNT = Font(name='Arial', size=10)

def analyze(file1_bytes, file2_bytes):
    df1 = pd.read_excel(io.BytesIO(file1_bytes), sheet_name=SHEET)
    df2 = pd.read_excel(io.BytesIO(file2_bytes), sheet_name=SHEET)

    for df in [df1, df2]:
        df[COL_PREMIUM] = pd.to_numeric(df[COL_PREMIUM], errors='coerce')
        df[COL_POLICY]  = df[COL_POLICY].astype(str).str.strip()
        df[COL_ID]      = df[COL_ID].astype(str).str.strip()

    df1d = df1.sort_values(COL_PREMIUM, ascending=False).drop_duplicates(COL_POLICY)
    df2d = df2.sort_values(COL_PREMIUM, ascending=False).drop_duplicates(COL_POLICY)

    keep = [COL_POLICY, COL_ID, COL_FNAME, COL_LNAME, COL_MFG, COL_PRODUCT, COL_PREMIUM, COL_AGENT]
    merged = df1d[keep].merge(df2d[keep], on=COL_POLICY, suffixes=('_jan', '_feb'))
    merged['שם לקוח']      = merged[COL_FNAME+'_feb'].fillna(merged[COL_FNAME+'_jan']) + ' ' + merged[COL_LNAME+'_feb'].fillna(merged[COL_LNAME+'_jan'])
    merged['ת.ז']           = merged[COL_ID+'_feb'].fillna(merged[COL_ID+'_jan'])
    merged['יצרן']          = merged[COL_MFG+'_feb'].fillna(merged[COL_MFG+'_jan'])
    merged['מוצר']          = merged[COL_PRODUCT+'_feb'].fillna(merged[COL_PRODUCT+'_jan'])
    merged['מת"ל']          = merged[COL_AGENT+'_feb'].fillna(merged[COL_AGENT+'_jan'])
    merged['פרמיה קודמת']  = merged[COL_PREMIUM+'_jan']
    merged['פרמיה נוכחית'] = merged[COL_PREMIUM+'_feb']
    merged['עלייה ₪']       = merged['פרמיה נוכחית'] - merged['פרמיה קודמת']
    merged['עלייה %']       = (merged['עלייה ₪'] / merged['פרמיה קודמת'].replace(0, float('nan'))) * 100

    result = merged[merged['עלייה %'] > THRESHOLD].sort_values('עלייה %', ascending=False).reset_index(drop=True)

    gone_pol = set(df1d[COL_POLICY]) - set(df2d[COL_POLICY])
    new_pol  = set(df2d[COL_POLICY]) - set(df1d[COL_POLICY])
    gone_df  = df1d[df1d[COL_POLICY].isin(gone_pol)]
    new_df   = df2d[df2d[COL_POLICY].isin(new_pol)]

    return merged, result, gone_df, new_df, df1d, df2d

# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(merged, result, gone_df, new_df, agent=None):
    wb = Workbook()
    HDR = PatternFill('solid', start_color='1F4E79')
    HF  = Font(name='Arial', bold=True, color='FFFFFF', size=11)
    RED = PatternFill('solid', start_color='FFCCCC')
    YEL = PatternFill('solid', start_color='FFF2CC')
    ALT = PatternFill('solid', start_color='EBF3FB')

    label = f' — {agent}' if agent else ''

    def hdr_row(ws, hdrs, wids, color='1F4E79'):
        fill = PatternFill('solid', start_color=color)
        for ci, (h, w) in enumerate(zip(hdrs, wids), 1):
            c = ws.cell(row=1, column=ci, value=h)
            c.font = HF; c.fill = fill; c.alignment = CTR; c.border = BORD
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[1].height = 22
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = f'A1:{get_column_letter(len(hdrs))}1'
        ws.sheet_view.rightToLeft = True

    # Sheet 1 — increase
    ws1 = wb.active
    ws1.title = 'עלייה בפרמיה >15%'
    hdrs1 = ['ת.ז','שם לקוח',"מס' פוליסה",'יצרן','מוצר','מת"ל','פרמיה קודמת (₪)','פרמיה נוכחית (₪)','עלייה (₪)','עלייה (%)']
    wids1 = [14,24,14,26,26,18,18,20,14,12]
    fmts1 = [None,None,None,None,None,None,'#,##0.00','#,##0.00','#,##0.00','0.0%']
    hdr_row(ws1, hdrs1, wids1)
    df_src = result[result['מת"ל'] == agent] if agent else result
    for ri, (_, row) in enumerate(df_src.iterrows()):
        pct  = row['עלייה %']
        fill = RED if pct > 30 else (YEL if ri % 2 == 0 else None)
        vals = [row['ת.ז'], row['שם לקוח'], row[COL_POLICY], row['יצרן'], row['מוצר'], row['מת"ל'],
                row['פרמיה קודמת'], row['פרמיה נוכחית'], row['עלייה ₪'], pct/100]
        for ci, (val, fmt) in enumerate(zip(vals, fmts1), 1):
            c = ws1.cell(row=ri+2, column=ci, value=val)
            c.font = DFNT; c.border = BORD
            c.alignment = CTR if ci in [1,3,10] else RGT
            if fill: c.fill = fill
            if fmt:  c.number_format = fmt

    # Sheet 2 — summary
    ws2 = wb.create_sheet('סיכום')
    ws2.sheet_view.rightToLeft = True
    ws2.column_dimensions['A'].width = 36
    ws2.column_dimensions['B'].width = 24
    TF = Font(name='Arial', bold=True, size=14, color='1F4E79')
    SF = Font(name='Arial', bold=True, size=11, color='1F4E79')

    t = ws2.cell(row=1, column=1, value=f'דוח עלייה בפרמיה{label}')
    t.font = TF; t.alignment = RGT; ws2.merge_cells('A1:B1')
    ws2.cell(row=2, column=1, value=f'הופק: {datetime.now().strftime("%d/%m/%Y")}').font = Font(name='Arial', size=10, color='666666')
    ws2.cell(row=2, column=1).alignment = RGT

    agent_data = merged[merged['מת"ל'] == agent] if agent else merged
    agent_result = result[result['מת"ל'] == agent] if agent else result

    def sr(ws, r, lbl, val, fmt=None, bold_v=False):
        lc = ws.cell(row=r, column=1, value=lbl)
        lc.font = Font(name='Arial', bold=True, size=11); lc.alignment = RGT
        vc = ws.cell(row=r, column=2, value=val)
        vc.font = Font(name='Arial', bold=bold_v, size=11); vc.alignment = RGT
        if fmt: vc.number_format = fmt
        return r + 1

    r = 4
    ws2.cell(row=r, column=1, value='נתוני פוליסות').font = SF; r += 1
    r = sr(ws2, r, 'סה"כ פוליסות שהושוו', len(agent_data))
    r = sr(ws2, r, 'פוליסות עם עלייה >15%', len(agent_result), bold_v=True)
    if len(agent_result) > 0:
        r = sr(ws2, r, 'ממוצע עלייה', agent_result['עלייה %'].mean()/100, '0.0%')
        r = sr(ws2, r, 'עלייה מקסימלית', agent_result['עלייה %'].max()/100, '0.0%')
        r = sr(ws2, r, 'סך עלייה חודשית (₪)', agent_result['עלייה ₪'].sum(), '#,##0.00', True)
    r += 1
    ws2.cell(row=r, column=1, value='פוליסות שנסגרו / חדשות').font = SF; r += 1
    g = gone_df[gone_df[COL_AGENT] == agent] if agent else gone_df
    n = new_df[new_df[COL_AGENT] == agent] if agent else new_df
    r = sr(ws2, r, 'פוליסות שנסגרו', len(g))
    r = sr(ws2, r, 'פוליסות חדשות', len(n))

    # Sheet 3 & 4 — gone / new
    def policy_sheet(ws, title, df_sub, color):
        ws.sheet_view.rightToLeft = True
        t = ws.cell(row=1, column=1, value=title)
        t.font = Font(name='Arial', bold=True, size=12, color=color); t.alignment = RGT
        ws.merge_cells(f'A1:{get_column_letter(7)}1')
        hdrs = ['ת.ז','שם פרטי','שם משפחה',"מס' פוליסה",'יצרן','מוצר','פרמיה (₪)']
        wids = [14,14,14,14,28,28,14]
        for ci, (h, w) in enumerate(zip(hdrs, wids), 1):
            c = ws.cell(row=2, column=ci, value=h)
            c.font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
            c.fill = PatternFill('solid', start_color=color)
            c.alignment = CTR; c.border = BORD
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = 'A3'
        alt = PatternFill('solid', start_color='FFE8E8' if color == '7B0000' else 'E8F5E8')
        for ri, (_, row) in enumerate(df_sub.iterrows()):
            fill = alt if ri % 2 == 0 else None
            vals = [row.get(COL_ID,''), row.get(COL_FNAME,''), row.get(COL_LNAME,''),
                    row.get(COL_POLICY,''), row.get(COL_MFG,''), row.get(COL_PRODUCT,''), row.get(COL_PREMIUM,0)]
            for ci, (val, fmt) in enumerate(zip(vals, [None]*6+['#,##0.00']), 1):
                c = ws.cell(row=ri+3, column=ci, value=val)
                c.font = DFNT; c.border = BORD; c.alignment = CTR if ci<=4 else RGT
                if fill: c.fill = fill
                if fmt:  c.number_format = fmt

    ws3 = wb.create_sheet('פוליסות שנסגרו')
    policy_sheet(ws3, f'פוליסות שנסגרו{label}', g, '7B0000')
    ws4 = wb.create_sheet('פוליסות חדשות')
    policy_sheet(ws4, f'פוליסות חדשות{label}', n, '1A5C1A')

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ── PDF builder ───────────────────────────────────────────────────────────────
def build_pdf(merged, result, gone_df, new_df, month_label, agent=None):
    _register_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            rightMargin=1.5*cm, leftMargin=1.5*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    title_s = ParagraphStyle('t', fontName=BASE_FONT, fontSize=16,
                              textColor=colors.HexColor('#1F4E79'), spaceAfter=4, alignment=2)
    sub_s   = ParagraphStyle('s', fontName=BASE_FONT, fontSize=10,
                              textColor=colors.grey, spaceAfter=8, alignment=2)
    sec_s   = ParagraphStyle('h', fontName=BASE_FONT, fontSize=12,
                              textColor=colors.HexColor('#1F4E79'), spaceBefore=10, spaceAfter=6, alignment=2)

    agent_data   = merged[merged['מת"ל'] == agent] if agent else merged
    agent_result = result[result['מת"ל'] == agent] if agent else result
    g = gone_df[gone_df[COL_AGENT] == agent] if agent else gone_df
    n = new_df[new_df[COL_AGENT] == agent] if agent else new_df

    story = []
    story.append(Paragraph(rh('דוח עלייה בפרמיה חודשית'), title_s))
    if agent:
        story.append(Paragraph(rh(f'סוכן: {agent}'), sub_s))
    story.append(Paragraph(rh(month_label), sub_s))
    story.append(Paragraph(rh(f'הופק: {datetime.now().strftime("%d/%m/%Y")}'), sub_s))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph(rh('סיכום'), sec_s))
    sum_data = [[rh('נושא'), rh('ערך')],
                [rh('פוליסות שהושוו'), str(len(agent_data))],
                [rh('פוליסות עם עלייה >15%'), str(len(agent_result))]]
    if len(agent_result) > 0:
        sum_data += [
            [rh('ממוצע עלייה'), f"{agent_result['עלייה %'].mean():.1f}%"],
            [rh('עלייה מקסימלית'), f"{agent_result['עלייה %'].max():.1f}%"],
            [rh('סך עלייה חודשית'), f"₪{agent_result['עלייה ₪'].sum():,.0f}"],
        ]
    sum_data += [[rh('פוליסות שנסגרו'), str(len(g))],
                 [rh('פוליסות חדשות'), str(len(n))]]

    st_tbl = Table(sum_data, colWidths=[10*cm, 6*cm])
    st_tbl.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1F4E79')),
        ('TEXTCOLOR', (0,0),(-1,0),colors.white),
        ('FONTNAME',  (0,0),(-1,-1),BASE_FONT),
        ('FONTSIZE',  (0,0),(-1,0),11),('FONTSIZE',(0,1),(-1,-1),10),
        ('ALIGN',     (0,0),(-1,-1),'RIGHT'),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#EBF3FB')]),
        ('GRID',      (0,0),(-1,-1),0.5,colors.HexColor('#CCCCCC')),
        ('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),
    ]))
    story.append(st_tbl)
    story.append(Spacer(1, 0.5*cm))

    if len(agent_result) > 0:
        story.append(Paragraph(rh(f'פוליסות עם עלייה >15% ({len(agent_result)})'), sec_s))
        th = [rh('ת.ז'),rh('שם לקוח'),rh("מס' פוליסה"),rh('יצרן'),
              rh('פרמיה קודמת'),rh('פרמיה נוכחית'),rh('עלייה ₪'),rh('עלייה %')]
        td = [th]
        for _, row in agent_result.iterrows():
            td.append([rh(row['ת.ז']),rh(row['שם לקוח']),rh(row[COL_POLICY]),rh(row['יצרן']),
                       f"₪{row['פרמיה קודמת']:,.0f}",f"₪{row['פרמיה נוכחית']:,.0f}",
                       f"₪{row['עלייה ₪']:,.0f}",f"{row['עלייה %']:.1f}%"])
        mt = Table(td, colWidths=[2.2*cm,3.5*cm,2.5*cm,3.8*cm,2.3*cm,2.5*cm,2.3*cm,1.8*cm], repeatRows=1)
        ts = [('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1F4E79')),
              ('TEXTCOLOR', (0,0),(-1,0),colors.white),
              ('FONTNAME',  (0,0),(-1,-1),BASE_FONT),
              ('FONTSIZE',  (0,0),(-1,0),9),('FONTSIZE',(0,1),(-1,-1),8),
              ('ALIGN',     (0,0),(-1,-1),'RIGHT'),
              ('GRID',      (0,0),(-1,-1),0.3,colors.HexColor('#CCCCCC')),
              ('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4)]
        for i,(_, row) in enumerate(agent_result.iterrows(),1):
            c = colors.HexColor('#FFCCCC') if row['עלייה %']>30 else (colors.HexColor('#FFF2CC') if i%2==0 else colors.white)
            ts.append(('BACKGROUND',(0,i),(-1,i),c))
        mt.setStyle(TableStyle(ts))
        story.append(mt)

    doc.build(story)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════
st.title("📊 מערכת דוחות פרמיה")
st.markdown("העלה את שני דוחות ה-Excel מ-Surense וקבל דוחות מוכנים להורדה.")
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown("**📁 דוח החודש הקודם**")
    f1 = st.file_uploader("", type=['xlsx'], key='f1', label_visibility='collapsed')
with col2:
    st.markdown("**📁 דוח החודש הנוכחי**")
    f2 = st.file_uploader("", type=['xlsx'], key='f2', label_visibility='collapsed')

if f1 and f2:
    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
    with st.spinner("מנתח נתונים..."):
        try:
            merged, result, gone_df, new_df, df1d, df2d = analyze(f1.read(), f2.read())
            agents = sorted(merged['מת"ל'].dropna().unique().tolist())
            month_label = f'{f1.name[:10]} ← {f2.name[:10]}'
        except Exception as e:
            st.error(f"שגיאה בקריאת הקבצים: {e}")
            st.stop()

    # ── Summary metrics ──
    st.subheader("📈 סיכום")
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="metric-box"><div class="val">{len(merged)}</div><div class="lbl">פוליסות שהושוו</div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="metric-box"><div class="val red">{len(result)}</div><div class="lbl">עלייה >15%</div></div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="metric-box"><div class="val">{len(gone_df)}</div><div class="lbl">פוליסות שנסגרו</div></div>', unsafe_allow_html=True)
    c4.markdown(f'<div class="metric-box"><div class="val green">{len(new_df)}</div><div class="lbl">פוליסות חדשות</div></div>', unsafe_allow_html=True)

    if len(result) > 0:
        st.markdown(f"""
        <br>
        <div style="background:#fff3cd;border-radius:8px;padding:12px 16px;border-right:4px solid #FF9500;text-align:right">
        ⚠️ <b>סך עלייה חודשית: ₪{result['עלייה ₪'].sum():,.0f}</b> &nbsp;|&nbsp;
        ממוצע עלייה: {result['עלייה %'].mean():.1f}% &nbsp;|&nbsp;
        מקסימום: {result['עלייה %'].max():.1f}%
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

    # ── Preview table ──
    if len(result) > 0:
        st.subheader("🔍 פוליסות עם עלייה >15%")
        preview = result[['שם לקוח','מת"ל','יצרן','פרמיה קודמת','פרמיה נוכחית','עלייה ₪','עלייה %']].copy()
        preview['עלייה %'] = preview['עלייה %'].map(lambda x: f"{x:.1f}%")
        preview['פרמיה קודמת'] = preview['פרמיה קודמת'].map(lambda x: f"₪{x:,.0f}")
        preview['פרמיה נוכחית'] = preview['פרמיה נוכחית'].map(lambda x: f"₪{x:,.0f}")
        preview['עלייה ₪'] = preview['עלייה ₪'].map(lambda x: f"₪{x:,.0f}")
        st.dataframe(preview, use_container_width=True, hide_index=True)

    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

    # ── Downloads ──
    st.subheader("📥 הורדת דוחות")

    # Combined
    with st.spinner("בונה דוח כולל..."):
        xl_all  = build_excel(merged, result, gone_df, new_df)
        pdf_all = build_pdf(merged, result, gone_df, new_df, month_label)

    st.markdown("**דוח כולל — כל הסוכנים**")
    ca, cb = st.columns(2)
    with ca:
        st.download_button("⬇️ הורד Excel", xl_all,
                           file_name=f"דוח_פרמיה_{datetime.now().strftime('%m_%Y')}.xlsx",
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    with cb:
        st.download_button("⬇️ הורד PDF", pdf_all,
                           file_name=f"דוח_פרמיה_{datetime.now().strftime('%m_%Y')}.pdf",
                           mime='application/pdf')

    # Per agent
    st.markdown("**דוחות לפי סוכן מטפל**")
    for agent in agents:
        with st.spinner(f"בונה דוח עבור {agent}..."):
            xl_a  = build_excel(merged, result, gone_df, new_df, agent=agent)
            pdf_a = build_pdf(merged, result, gone_df, new_df, month_label, agent=agent)
        safe = agent.replace(' ','_')
        st.markdown(f"**{agent}**")
        da, db = st.columns(2)
        with da:
            st.download_button(f"⬇️ Excel — {agent}", xl_a,
                               file_name=f"דוח_פרמיה_{safe}_{datetime.now().strftime('%m_%Y')}.xlsx",
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               key=f'xl_{safe}')
        with db:
            st.download_button(f"⬇️ PDF — {agent}", pdf_a,
                               file_name=f"דוח_פרמיה_{safe}_{datetime.now().strftime('%m_%Y')}.pdf",
                               mime='application/pdf',
                               key=f'pdf_{safe}')

else:
    st.info("⬆️ העלה את שני קבצי ה-Excel כדי להתחיל")
