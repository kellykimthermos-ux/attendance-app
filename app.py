import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import timedelta
from io import BytesIO
import math

st.set_page_config(page_title="선택적 근무 · 초과시간 관리", page_icon="⏱", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&display=swap');
html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; }
.app-header {
    background: linear-gradient(135deg, #1F3864 0%, #2F5496 100%);
    border-radius: 12px; padding: 28px 32px; margin-bottom: 28px; color: white;
}
.app-header h1 { font-size: 22px; font-weight: 700; margin: 0 0 6px 0; }
.app-header p  { font-size: 13px; margin: 0; opacity: 0.75; }
.rule-box {
    background: #F0F4FB; border-left: 4px solid #2F5496;
    border-radius: 0 8px 8px 0; padding: 14px 18px;
    font-size: 13px; color: #2C3E50; line-height: 1.8; margin-bottom: 24px;
}
.card-wrap { display: flex; gap: 14px; flex-wrap: wrap; margin-bottom: 24px; }
.summary-card {
    background: white; border: 1px solid #E0E8F5; border-radius: 10px;
    padding: 16px 22px; min-width: 160px; flex: 1;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
.summary-card .label { font-size: 11px; color: #7F8C8D; font-weight: 500; margin-bottom: 6px; }
.summary-card .value { font-size: 22px; font-weight: 700; color: #1F3864; }
.summary-card .sub   { font-size: 11px; color: #95A5A6; margin-top: 2px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="app-header">
    <h1>⏱ 선택적 근무시간제 · 초과시간 관리</h1>
    <p>인트라넷 근태현황 원본 파일을 업로드하면 인정근무시간과 잔여시간을 자동 계산합니다</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="rule-box">
    📌 <strong>계산 기준</strong> &nbsp;|&nbsp;
    인정출근 = 실제 출근시각 <strong>30분 올림</strong> (08:00 이전은 08:00 고정) &nbsp;·&nbsp;
    인정퇴근 = 실제 퇴근시각 <strong>30분 내림</strong> &nbsp;·&nbsp;
    인정근무 = 체류시간 − 점심 1H &nbsp;·&nbsp;
    8H 초과 → <strong>+적립</strong> / 8H 미달 → <strong>−차감</strong> (30분 단위)
</div>
""", unsafe_allow_html=True)

# ── 헬퍼 함수 ──────────────────────────────────────────────────
def parse_hms(s):
    if pd.isna(s) or str(s).strip() in ('', 'False', 'NaT'):
        return None
    try:
        p = str(s).strip().split(':')
        return timedelta(hours=int(p[0]), minutes=int(p[1]), seconds=int(p[2]))
    except:
        return None

def ceil30(td):
    m = int(td.total_seconds() // 60)
    return timedelta(minutes=math.ceil(m / 30) * 30)

def floor30(td):
    m = int(td.total_seconds() // 60)
    return timedelta(minutes=(m // 30) * 30)

def fmt_td(td):
    if td is None: return ''
    m = int(td.total_seconds() // 60)
    return f"{m // 60}:{m % 60:02d}"

def fmt_net(minutes):
    h = abs(minutes) // 60
    m = abs(minutes) % 60
    if minutes > 0:   return f"+{h}:{m:02d}"
    elif minutes < 0: return f"-{h}:{m:02d}"
    else:             return "0:00"

EIGHT_AM = timedelta(hours=8)
LUNCH    = timedelta(hours=1)
EIGHT_H  = timedelta(hours=8)

# ── 계산 로직 ──────────────────────────────────────────────────
def process(df_raw):
    detail_rows = []

    for _, r in df_raw.iterrows():
        date_str = str(r['일자'])
        emp_name = str(r['이름'])
        emp_rank = str(r['직위'])
        in_r     = parse_hms(r['출근시간'])
        out_r    = parse_hms(r['퇴근시간'])
        vac      = parse_hms(r['총휴가시간'])

        if not emp_name or emp_name in ('nan', '이름', ''):
            continue

        # 휴무/휴가
        if in_r is None or out_r is None:
            note = '휴무/휴가' if (vac and vac.total_seconds() > 0) else '휴무'
            detail_rows.append({
                '일자': date_str, '이름': emp_name, '직위': emp_rank,
                '실출근': '', '실퇴근': '',
                '인정출근': '', '인정퇴근': '',
                '인정근무시간': '', '8H 대비': '', 'net_min': 0, '비고': note
            })
            continue

        # 인정 출퇴근
        adj_in  = EIGHT_AM if in_r < EIGHT_AM else ceil30(in_r)
        adj_out = floor30(out_r)
        stay    = max(adj_out - adj_in, timedelta(0))
        work    = max(stay - LUNCH, timedelta(0))

        # 8H 대비 차이 → 30분 단위 절사
        diff_min = int((work - EIGHT_H).total_seconds() // 60)
        net_min  = (abs(diff_min) // 30) * 30
        net_min  = net_min if diff_min >= 0 else -net_min

        note = '08:00 이전→보정' if in_r < EIGHT_AM else ''

        detail_rows.append({
            '일자': date_str, '이름': emp_name, '직위': emp_rank,
            '실출근': str(r['출근시간']), '실퇴근': str(r['퇴근시간']),
            '인정출근': fmt_td(adj_in), '인정퇴근': fmt_td(adj_out),
            '인정근무시간': fmt_td(work), '8H 대비': fmt_net(net_min),
            'net_min': net_min, '비고': note
        })

    detail = pd.DataFrame(detail_rows)

    # 요약
    summary_rows = []
    for name in detail['이름'].unique():
        g = detail[detail['이름'] == name]
        rank_vals = g[g['직위'].notna() & (g['직위'] != 'nan')]['직위']
        rank      = rank_vals.iloc[0] if len(rank_vals) > 0 else ''
        total_min = int(g['net_min'].sum())
        blocks    = total_min // 30

        if blocks > 0:
            bigo = f"{blocks}회 × 30분 조기퇴근 사용 가능"
        elif blocks < 0:
            bigo = f"{abs(blocks)}회 × 30분 추가 근무 필요"
        else:
            bigo = '-'

        summary_rows.append({
            '이름': name, '직위': rank,
            '누적 잔여시간': fmt_net(total_min),
            '사용가능 블록(30분)': blocks,
            '비고': bigo
        })

    return detail, pd.DataFrame(summary_rows)

# ── Excel 출력 ─────────────────────────────────────────────────
def to_excel(detail, summary):
    wb  = Workbook()
    C_BLU = "2F5496"; C_NAV = "1F3864"
    C_YEL = "FFF2CC"; C_PNK = "FCE4EC"; C_GRN = "E2EFDA"; C_BRD = "B8CCE4"
    thin = Side(style='thin', color=C_BRD)
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hc(ws, row, col, val, bg=C_NAV, fg="FFFFFF", sz=10, bold=True):
        c = ws.cell(row=row, column=col, value=val)
        c.font      = Font(name='Arial', bold=bold, color=fg, size=sz)
        c.fill      = PatternFill('solid', fgColor=bg)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border    = brd
        return c

    def dc(ws, row, col, val, bg=None, bold=False, align='center', sz=9, fg="000000"):
        c = ws.cell(row=row, column=col, value=val)
        c.font      = Font(name='Arial', bold=bold, size=sz, color=fg)
        c.alignment = Alignment(horizontal=align, vertical='center')
        c.border    = brd
        if bg: c.fill = PatternFill('solid', fgColor=bg)
        return c

    # ── Sheet1: 누적초과 요약 ─────────────────────────────────
    ws1 = wb.active; ws1.title = "누적초과 요약"
    ws1.sheet_view.showGridLines = False
    ws1.row_dimensions[1].height = 8
    ws1.column_dimensions['A'].width = 3

    ws1.merge_cells('B2:G2')
    c = ws1['B2']; c.value = "선택적 근무시간제 · 잔여시간 현황"
    c.font = Font(name='Arial', bold=True, size=14, color=C_BLU)
    c.alignment = Alignment(horizontal='left', vertical='center')
    ws1.row_dimensions[2].height = 28

    ws1.merge_cells('B3:G3')
    c = ws1['B3']; c.value = "※ 인정출근 30분 올림(08:00 이전→08:00) | 인정퇴근 30분 내림 | 점심 1H 공제 | 8H 초과→+적립 / 미달→−차감 (30분 단위)"
    c.font = Font(name='Arial', size=8, color="595959")
    c.alignment = Alignment(horizontal='left', vertical='center')
    ws1.row_dimensions[3].height = 15
    ws1.row_dimensions[4].height = 6
    ws1.row_dimensions[5].height = 32

    for col, h, w in zip([2,3,4,5,6], ['이름','직위','누적 잔여시간','사용가능 블록(30분)','비고'], [14,10,16,20,34]):
        hc(ws1, 5, col, h)
        ws1.column_dimensions[get_column_letter(col)].width = w

    for i, row in summary.iterrows():
        r   = i + 6
        bg  = C_GRN if i % 2 == 0 else None
        blk = row['사용가능 블록(30분)']
        ws1.row_dimensions[r].height = 24
        dc(ws1, r, 2, row['이름'],  bg=bg, bold=True, sz=10)
        dc(ws1, r, 3, row['직위'],  bg=bg, sz=9)
        if blk > 0:
            dc(ws1, r, 4, row['누적 잔여시간'], bg=C_YEL, bold=True, sz=11, fg="7F4F00")
            dc(ws1, r, 5, blk, bg=C_YEL, bold=True, sz=10, fg="7F4F00")
        elif blk < 0:
            dc(ws1, r, 4, row['누적 잔여시간'], bg=C_PNK, bold=True, sz=11, fg="8B0000")
            dc(ws1, r, 5, blk, bg=C_PNK, bold=True, sz=10, fg="8B0000")
        else:
            dc(ws1, r, 4, row['누적 잔여시간'], bg=bg, sz=10)
            dc(ws1, r, 5, blk, bg=bg, sz=10)
        dc(ws1, r, 6, row['비고'], bg=bg, sz=9, align='left')
    ws1.freeze_panes = 'B6'

    # ── Sheet2: 일별 상세 ─────────────────────────────────────
    ws2 = wb.create_sheet("일별 상세")
    ws2.sheet_view.showGridLines = False
    ws2.row_dimensions[1].height = 8
    ws2.column_dimensions['A'].width = 3
    ws2.column_dimensions['L'].width = 3

    ws2.merge_cells('B2:K2')
    c = ws2['B2']; c.value = "선택적 근무시간제 · 일별 인정근무시간 상세"
    c.font = Font(name='Arial', bold=True, size=13, color=C_BLU)
    c.alignment = Alignment(horizontal='left', vertical='center')
    ws2.row_dimensions[2].height = 26

    ws2.merge_cells('B3:K3')
    c = ws2['B3']; c.value = "※ 인정근무 = (인정퇴근−인정출근)−점심1H | +초과(노랑) / −미달(분홍) 표시 | 30분 단위 절사"
    c.font = Font(name='Arial', size=8, color="595959")
    c.alignment = Alignment(horizontal='left', vertical='center')
    ws2.row_dimensions[3].height = 15
    ws2.row_dimensions[4].height = 6
    ws2.row_dimensions[5].height = 32

    hdrs = ['일자','이름','직위','실출근','실퇴근','인정출근','인정퇴근','인정근무시간','8H 대비','비고']
    wids = [12, 12, 8, 11, 11, 11, 11, 14, 14, 18]
    for col, h, w in zip(range(2, 12), hdrs, wids):
        hc(ws2, 5, col, h)
        ws2.column_dimensions[get_column_letter(col)].width = w

    color_pool = ["EBF3FB","FFF9F0","F0FBF0","FDF0FB","FFF5E6","F5F0FF"]
    person_colors = {name: color_pool[i % len(color_pool)]
                     for i, name in enumerate(detail['이름'].unique())}

    for i, row in detail.iterrows():
        r        = i + 6
        base_bg  = person_colors.get(row['이름'], "FFFFFF")
        net      = row['net_min']
        ws2.row_dimensions[r].height = 18

        vals = [row['일자'], row['이름'], row['직위'],
                row['실출근'], row['실퇴근'],
                row['인정출근'], row['인정퇴근'],
                row['인정근무시간'], row['8H 대비'], row['비고']]

        for col_idx, val in enumerate(vals):
            col = col_idx + 2
            if col in [9, 10] and net > 0:   # 초과 → 노랑
                dc(ws2, r, col, val, bg=C_YEL, bold=True, sz=9, fg="7F4F00")
            elif col in [9, 10] and net < 0: # 미달 → 분홍
                dc(ws2, r, col, val, bg=C_PNK, bold=True, sz=9, fg="8B0000")
            elif col == 11 and '보정' in str(row['비고']):
                dc(ws2, r, col, val, bg=C_PNK, sz=8, align='left')
            elif col in [5, 6]:
                dc(ws2, r, col, val, bg="F5F5F5", sz=8)
            else:
                dc(ws2, r, col, val, bg=base_bg, sz=9)
    ws2.freeze_panes = 'B6'

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── 메인 UI ────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "인트라넷 근태현황 파일 업로드",
    type=['xlsx'],
    help="부서_기간별근태현황_XXXXXX.xlsx 원본 파일"
)
st.markdown('<p style="font-size:12px;color:#95A5A6;">📂 인트라넷 다운로드 원본 그대로 업로드 (가공 불필요)</p>', unsafe_allow_html=True)

if uploaded:
    with st.spinner("계산 중..."):
        try:
            df_all = pd.read_excel(uploaded, sheet_name='P_써모스코리아 기간별 근태현황', header=None)

            # 조회기간 추출
            try:    period = str(df_all.iloc[4, 2])
            except: period = ''

            # 원본 컬럼: B=1, D=3, E=4, H=7, K=10, U=20 (0-indexed)
            df_data = df_all.iloc[7:, [1, 3, 4, 7, 10, 20]].copy()
            df_data.columns = ['일자', '이름', '직위', '출근시간', '퇴근시간', '총휴가시간']
            df_data = df_data[df_data['일자'].notna()].reset_index(drop=True)

            detail, summary = process(df_data)

            # 지표
            total_people        = len(summary)
            people_with_surplus = len(summary[summary['사용가능 블록(30분)'] > 0])
            people_with_deficit = len(summary[summary['사용가능 블록(30분)'] < 0])
            best_idx            = summary['사용가능 블록(30분)'].values.argmax()
            max_val             = summary['누적 잔여시간'].iloc[best_idx]
            max_name            = summary['이름'].iloc[best_idx]

            st.markdown(f"""
            <div class="card-wrap">
                <div class="summary-card">
                    <div class="label">조회기간</div>
                    <div class="value" style="font-size:15px;">{period}</div>
                </div>
                <div class="summary-card">
                    <div class="label">대상 인원</div>
                    <div class="value">{total_people}명</div>
                </div>
                <div class="summary-card">
                    <div class="label">잔여시간 보유</div>
                    <div class="value" style="color:#1F6B2F;">{people_with_surplus}명</div>
                    <div class="sub">초과 적립자</div>
                </div>
                <div class="summary-card">
                    <div class="label">차감 발생</div>
                    <div class="value" style="color:#C0392B;">{people_with_deficit}명</div>
                    <div class="sub">8H 미달 차감자</div>
                </div>
                <div class="summary-card">
                    <div class="label">최대 누적</div>
                    <div class="value">{max_val}</div>
                    <div class="sub">{max_name}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            tab1, tab2 = st.tabs(["📊 누적잔여 요약", "📋 일별 상세"])

            with tab1:
                st.dataframe(
                    summary,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "이름":             st.column_config.TextColumn("이름", width=100),
                        "직위":             st.column_config.TextColumn("직위", width=80),
                        "누적 잔여시간":    st.column_config.TextColumn("누적 잔여시간", width=120),
                        "사용가능 블록(30분)": st.column_config.NumberColumn("사용가능 블록(30분)", width=160),
                        "비고":             st.column_config.TextColumn("비고", width=280),
                    }
                )

            with tab2:
                col1, _ = st.columns([2, 5])
                with col1:
                    names    = ['전체'] + list(detail['이름'].unique())
                    selected = st.selectbox("팀원 선택", names)

                df_show = detail if selected == '전체' else detail[detail['이름'] == selected]
                df_show = df_show.drop(columns=['net_min'], errors='ignore')

                st.dataframe(
                    df_show,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "일자":         st.column_config.TextColumn("일자", width=100),
                        "이름":         st.column_config.TextColumn("이름", width=90),
                        "직위":         st.column_config.TextColumn("직위", width=70),
                        "실출근":       st.column_config.TextColumn("실출근", width=90),
                        "실퇴근":       st.column_config.TextColumn("실퇴근", width=90),
                        "인정출근":     st.column_config.TextColumn("인정출근", width=90),
                        "인정퇴근":     st.column_config.TextColumn("인정퇴근", width=90),
                        "인정근무시간": st.column_config.TextColumn("인정근무시간", width=110),
                        "8H 대비":      st.column_config.TextColumn("8H 대비(+초과/−미달)", width=150),
                        "비고":         st.column_config.TextColumn("비고", width=140),
                    }
                )

            st.divider()
            fname      = f"선택적근무_잔여시간관리_{period.replace(' ','').replace('~','_')}.xlsx" if period else "선택적근무_잔여시간관리.xlsx"
            excel_buf  = to_excel(detail, summary)
            st.download_button(
                label="⬇️ Excel 다운로드",
                data=excel_buf,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"파일 처리 중 오류가 발생했습니다: {e}")
            st.info("'P_써모스코리아 기간별 근태현황' 시트가 포함된 인트라넷 원본 파일인지 확인해주세요.")
else:
    st.info("⬆️ 위에서 파일을 업로드하면 자동으로 계산됩니다.")
