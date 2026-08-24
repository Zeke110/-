import datetime
import io
import re
import urllib.parse
from collections import defaultdict

import gspread
import openpyxl
import pandas as pd

# pandas 최신 버전(3.0+)의 "future.infer_string" 옵션이 켜져 있으면, 빈 문자열만
# 있는 컬럼에 나중에 숫자를 넣으려 할 때 dtype 오류가 나서 꺼둔다.
# (이 문제로 병합 후 "수량" 값이 사라지는 버그가 있었음 — 아래에서 수정)
pd.set_option("future.infer_string", False)

import requests
import streamlit as st
import streamlit.components.v1 as components
from google.oauth2.service_account import Credentials

# ============================================================================
# 시약 & 용액 재고 관리 (Streamlit 웹앱 버전)
#
# 기존 Tkinter 데스크톱 프로그램(lab_chemical_manager.py)의 로직을 그대로
# 옮긴 웹앱입니다. PC/모바일 브라우저에서 동일하게 열람·수정할 수 있습니다.
#
# Google Sheets를 "공유 저장소"로 연결하면 여러 사람이 같은 데이터를
# 보고 저장할 수 있습니다 (설정 방법은 README.md 참고).
# ============================================================================

GSHEET_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

st.set_page_config(page_title="시약 & 용액 재고 관리", layout="wide")

COLUMNS = [
    "제품명",
    "제조사명",
    "CAT No.",
    "CAS No.",
    "용량",
    "수량",
    "미개봉",
    "개봉",
    "보관위치",
    "개봉 시 유통기한",
    "유해·위험성",
    "등록일",
]

CAPACITY_UNITS = ["g", "kg", "mL", "L"]

STORAGE_LOCATIONS = [
    "환기시약장",
    "위험물 보관함",
    "2 시약장",
    "데시케이터 1",
    "데시케이터 2",
    "냉장고",
]

# ----------------------------------------------------------------------------
# 실험실 배치도 (보관위치 지도)
#
# 각 보관위치 이름을 배치도 위의 사각형 좌표(x, y, width, height)에 매핑해서,
# 보관위치를 선택하면 그 구역만 빨간 테두리로 강조되도록 한다.
# ----------------------------------------------------------------------------
FLOOR_PLAN_RECTS = {
    "환기시약장": (568, 348, 72, 60),
    "위험물 보관함": (568, 237, 72, 60),
    "2 시약장": (397, 365, 52, 73),  # 배치도 상의 "2 보관함" 칸
    "데시케이터 1": (163, 178, 69, 59),
    "데시케이터 2": (377, 710, 67, 68),
    "냉장고": (295, 710, 68, 68),
}


def build_floor_plan_svg(highlight_location=None):
    """실험실 배치도 SVG를 만든다. highlight_location이 주어지면 해당 구역
    사각형에 빨간 강조 테두리를 덧그린다."""

    highlight_rect = ""
    if highlight_location and highlight_location in FLOOR_PLAN_RECTS:
        hx, hy, hw, hh = FLOOR_PLAN_RECTS[highlight_location]
        pad = 5
        highlight_rect = (
            f'<rect x="{hx - pad}" y="{hy - pad}" width="{hw + pad * 2}" '
            f'height="{hh + pad * 2}" fill="none" stroke="#e21f1f" '
            f'stroke-width="5" rx="4"><animate attributeName="opacity" '
            f'values="1;0.35;1" dur="1.2s" repeatCount="indefinite"/></rect>'
        )

    return f"""
<svg width="100%" viewBox="0 0 900 800" xmlns="http://www.w3.org/2000/svg"
     style="font-family:sans-serif;max-width:820px;display:block;margin:0 auto;">
<style>.ts{{font-size:15px;}} .t{{font-size:17px;}}</style>

<path d="M40 46 H640 V660 H535 V778 H40 Z" fill="none" stroke="#111" stroke-width="3"/>

<g><rect x="40" y="46" width="72" height="119" fill="#f0d4ef" stroke="#333" stroke-width="1"/><text class="ts" x="76" y="105" text-anchor="middle" dominant-baseline="central">싱크대</text></g>
<g><rect x="112" y="46" width="162" height="72" fill="#4a9e3f" stroke="#333" stroke-width="1"/><text class="ts" x="193" y="82" text-anchor="middle" dominant-baseline="central" fill="#fff">후드</text></g>
<g><rect x="274" y="46" width="111" height="72" fill="#cdeeff" stroke="#333" stroke-width="1"/><text class="ts" x="329" y="82" text-anchor="middle" dominant-baseline="central">글로브 박스</text></g>
<g><rect x="385" y="46" width="68" height="72" fill="#9e9e9e" stroke="#333" stroke-width="1"/><text class="ts" x="419" y="82" text-anchor="middle" dominant-baseline="central" fill="#fff">이빔</text></g>
<g><rect x="506" y="46" width="62" height="72" fill="#9e9e9e" stroke="#333" stroke-width="1"/><text class="ts" x="537" y="82" text-anchor="middle" dominant-baseline="central" fill="#fff">프린터</text></g>
<g><rect x="568" y="46" width="72" height="191" fill="#4a9e3f" stroke="#333" stroke-width="1"/><text class="ts" x="604" y="141" text-anchor="middle" dominant-baseline="central" fill="#fff">후드</text></g>

<g><rect x="40" y="165" width="72" height="72" fill="#1b6a86" stroke="#333" stroke-width="1"/><text class="ts" x="76" y="201" text-anchor="middle" dominant-baseline="central" fill="#fff">책상 2</text></g>
<g><rect x="40" y="237" width="72" height="81" fill="#1b6a86" stroke="#333" stroke-width="1"/><text class="ts" x="76" y="277" text-anchor="middle" dominant-baseline="central" fill="#fff">책상 1</text></g>

<g><rect x="163" y="178" width="69" height="59" fill="#ffee33" stroke="#333" stroke-width="1"/><text class="ts" x="197" y="207" text-anchor="middle" dominant-baseline="central"><tspan x="197" dy="-6">데시케이터</tspan><tspan x="197" dy="14">1</tspan></text></g>
<g><rect x="232" y="178" width="267" height="59" fill="#1b6a86" stroke="#333" stroke-width="1"/><text class="ts" x="365" y="207" text-anchor="middle" dominant-baseline="central" fill="#fff">책상 4</text></g>
<g><rect x="163" y="237" width="69" height="128" fill="#1b6a86" stroke="#333" stroke-width="1"/><text class="ts" x="197" y="301" text-anchor="middle" dominant-baseline="central" fill="#fff">책상 3</text></g>
<g><rect x="163" y="365" width="69" height="73" fill="#9e9e9e" stroke="#333" stroke-width="1"/><text class="ts" x="197" y="392" text-anchor="middle" dominant-baseline="central" fill="#fff"><tspan x="197" dy="-6">Probe</tspan><tspan x="197" dy="14">station</tspan></text></g>

<g><rect x="316" y="301" width="69" height="201" fill="#1b6a86" stroke="#333" stroke-width="1"/><text class="ts" x="350" y="401" text-anchor="middle" dominant-baseline="central" fill="#fff">책상 5</text></g>
<g><rect x="385" y="301" width="68" height="64" fill="#1b6a86" stroke="#333" stroke-width="1"/><text class="ts" x="419" y="333" text-anchor="middle" dominant-baseline="central" fill="#fff">책상 6</text></g>
<g><rect x="397" y="365" width="52" height="73" fill="#ffee33" stroke="#333" stroke-width="1"/><text class="ts" x="423" y="401" text-anchor="middle" dominant-baseline="central"><tspan x="423" dy="-6">2</tspan><tspan x="423" dy="14">보관함</tspan></text></g>

<g><rect x="568" y="237" width="72" height="60" fill="#ffee33" stroke="#333" stroke-width="1"/><text class="ts" x="604" y="267" text-anchor="middle" dominant-baseline="central"><tspan x="604" dy="-6">위험물</tspan><tspan x="604" dy="14">보관함</tspan></text></g>
<g><rect x="568" y="348" width="72" height="60" fill="#ffee33" stroke="#333" stroke-width="1"/><text class="ts" x="604" y="378" text-anchor="middle" dominant-baseline="central"><tspan x="604" dy="-6">환기</tspan><tspan x="604" dy="14">시약장</tspan></text></g>
<g><rect x="568" y="408" width="72" height="115" fill="#cdeeff" stroke="#333" stroke-width="1"/><text class="ts" x="604" y="465" text-anchor="middle" dominant-baseline="central"><tspan x="604" dy="-6">스핀</tspan><tspan x="604" dy="14">코터</tspan></text></g>

<g><rect x="40" y="510" width="72" height="149" fill="#b2e2b2" stroke="#333" stroke-width="1"/><text class="ts" x="76" y="585" text-anchor="middle" dominant-baseline="central">선반 5</text></g>

<g><rect x="163" y="578" width="251" height="55" fill="#9e9e9e" stroke="#333" stroke-width="1"/><text class="ts" x="288" y="605" text-anchor="middle" dominant-baseline="central" fill="#fff">Dry-spinning</text></g>

<g><rect x="453" y="578" width="76" height="200" fill="#4a9e3f" stroke="#333" stroke-width="1"/><text class="ts" x="491" y="678" text-anchor="middle" dominant-baseline="central" fill="#fff">후드</text></g>
<g><rect x="529" y="578" width="111" height="81" fill="#b2e2b2" stroke="#333" stroke-width="1"/><text class="ts" x="584" y="618" text-anchor="middle" dominant-baseline="central">입구 선반</text></g>

<g><rect x="163" y="710" width="132" height="68" fill="#f0d4ef" stroke="#333" stroke-width="1"/><text class="ts" x="229" y="744" text-anchor="middle" dominant-baseline="central">싱크대</text></g>
<g><rect x="295" y="710" width="68" height="68" fill="#1b6a86" stroke="#333" stroke-width="1"/><text class="ts" x="329" y="744" text-anchor="middle" dominant-baseline="central" fill="#fff">냉장고</text></g>
<g><rect x="377" y="710" width="67" height="68" fill="#ffee33" stroke="#333" stroke-width="1"/><text class="ts" x="410" y="744" text-anchor="middle" dominant-baseline="central"><tspan x="410" dy="-6">데시케이터</tspan><tspan x="410" dy="14">2</tspan></text></g>

<g>
  <rect x="720" y="180" width="42" height="42" fill="#1b6a86" stroke="#333" stroke-width="1"/><text class="t" x="774" y="201" dominant-baseline="central">책상</text>
  <rect x="720" y="240" width="42" height="42" fill="#b2e2b2" stroke="#333" stroke-width="1"/><text class="t" x="774" y="261" dominant-baseline="central">선반</text>
  <rect x="720" y="300" width="42" height="42" fill="#f0d4ef" stroke="#333" stroke-width="1"/><text class="t" x="774" y="321" dominant-baseline="central">싱크대</text>
  <rect x="720" y="360" width="42" height="42" fill="#4a9e3f" stroke="#333" stroke-width="1"/><text class="t" x="774" y="381" dominant-baseline="central">후드</text>
  <rect x="720" y="420" width="42" height="42" fill="#9e9e9e" stroke="#333" stroke-width="1"/><text class="t" x="774" y="441" dominant-baseline="central">장비</text>
  <rect x="720" y="480" width="42" height="42" fill="#ffee33" stroke="#333" stroke-width="1"/><text class="t" x="774" y="501" dominant-baseline="central">시약장</text>
</g>

{highlight_rect}
</svg>
"""


# ----------------------------------------------------------------------------
# 엑셀 읽기 / 헤더 자동 인식 (Tkinter 버전과 동일한 로직)
# ----------------------------------------------------------------------------
def find_header_row(file_bytes, columns, max_scan=20):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active

    best_row, best_score = 0, -1
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True)):
        row_vals = [str(c).strip() if c is not None else "" for c in row]
        score = sum(1 for col in columns if col in row_vals)
        if score > best_score:
            best_score, best_row = score, i
    wb.close()

    if best_score < max(2, len(columns) // 2):
        return 0
    return best_row


def read_excel_normalized(file_bytes, columns):
    """엑셀을 읽어 columns 기준으로 정리된 DataFrame과 누락 컬럼 목록 반환"""
    header_row = find_header_row(file_bytes, columns)
    df = pd.read_excel(io.BytesIO(file_bytes), header=header_row)

    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")

    missing_cols = [c for c in columns if c not in df.columns]
    for col in missing_cols:
        df[col] = ""

    df = df[columns].reset_index(drop=True)
    # 빈 문자열 컬럼이 고정 dtype으로 인식되어 이후 숫자 입력이 막히는 것을 방지
    df = df.astype(object).where(pd.notna(df), "")
    return df, missing_cols


def build_match_key(row):
    name = str(row.get("제품명", "")).strip().lower()
    cat = str(row.get("CAT No.", "")).strip()
    cas = str(row.get("CAS No.", "")).strip()

    if cat and cat.lower() != "nan":
        return (name, "cat", cat)
    if cas and cas.lower() != "nan":
        return (name, "cas", cas)
    return (name, "name", "")


def merge_dataframes(base_df, new_df, columns):
    """base_df에 new_df를 병합(동기화 방식).
    - 매칭되는 항목: 새 파일에 값이 있는 칸은 최신 값으로 덮어씀(값이 바뀐 경우
      갱신됨). 새 파일에 값이 없는 칸(예: 수동으로만 채운 미개봉/개봉/보관위치
      등)은 기존 값을 그대로 유지.
    - 새 파일에 없는 기존 항목(예: 수동으로 추가한 항목)은 건드리지 않고 유지.
    - 새 파일에만 있는 항목은 새 행으로 추가.
    """
    base_df = base_df.copy()

    key_to_idx = defaultdict(list)
    for idx, row in base_df.iterrows():
        key_to_idx[build_match_key(row)].append(idx)

    matched, updated_cells = 0, 0
    new_rows = []

    for _, srow in new_df.iterrows():
        key = build_match_key(srow)
        idxs = key_to_idx.get(key)
        if idxs:
            matched += 1
            for idx in idxs:
                for col in columns:
                    existing_val = base_df.at[idx, col]
                    new_val = srow[col]
                    new_has_val = pd.notna(new_val) and str(new_val).strip() != ""
                    existing_str = "" if pd.isna(existing_val) else str(existing_val).strip()
                    if new_has_val and str(new_val).strip() != existing_str:
                        base_df.at[idx, col] = new_val
                        updated_cells += 1
        else:
            new_rows.append(srow)

    if new_rows:
        base_df = pd.concat([base_df, pd.DataFrame(new_rows)], ignore_index=True)

    base_df = base_df.reset_index(drop=True)
    return base_df, matched, updated_cells, len(new_rows)


def df_to_excel_bytes(df):
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def get_gsheet_client():
    """secrets.toml에 gcp_service_account가 설정돼 있으면 gspread 클라이언트를,
    없으면 None을 반환 (Google Sheets 연동 없이도 엑셀 업/다운로드만으로 동작)"""
    try:
        if "gcp_service_account" not in st.secrets:
            return None
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=GSHEET_SCOPES)
        return gspread.authorize(creds)
    except Exception:
        return None


def open_worksheet(client, spreadsheet_id_or_url, worksheet_name):
    if spreadsheet_id_or_url.strip().startswith("http"):
        sh = client.open_by_url(spreadsheet_id_or_url.strip())
    else:
        sh = client.open_by_key(spreadsheet_id_or_url.strip())

    try:
        ws = sh.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(
            title=worksheet_name, rows=1000, cols=len(COLUMNS) + 2
        )
        ws.update([COLUMNS])
    return ws


def load_df_from_worksheet(ws):
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=COLUMNS)

    df = pd.DataFrame(records)
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in COLUMNS if c not in df.columns]
    for c in missing:
        df[c] = ""

    df = df[COLUMNS].reset_index(drop=True)
    df = df.astype(object).where(pd.notna(df), "")
    return df


def save_df_to_worksheet(ws, df):
    """시트 전체를 지우고 현재 데이터로 다시 씀 (단순 덮어쓰기 방식이라
    동시에 여러 명이 저장하면 나중에 저장한 내용이 우선됨)"""
    values = [COLUMNS] + df[COLUMNS].astype(str).values.tolist()
    ws.clear()
    ws.update(values)


def recompute_quantity(df):
    """미개봉/개봉에 값이 있을 때만 그 합으로 수량을 자동 계산.
    둘 다 비어있으면(예: 수량만 있는 엑셀을 불러온 경우) 기존 수량 값을 그대로 둔다."""
    def to_num(v):
        s = str(v).strip()
        if s == "" or s.lower() == "nan":
            return None
        try:
            return int(float(v))
        except (TypeError, ValueError):
            return None

    new_qty = []
    for u, o, existing_q in zip(df["미개봉"], df["개봉"], df["수량"]):
        u_n = to_num(u)
        o_n = to_num(o)
        if u_n is None and o_n is None:
            new_qty.append(existing_q)
        else:
            new_qty.append((u_n or 0) + (o_n or 0))
    # pandas 최신 버전에서 리스트를 그대로 대입하면 컬럼이 엄격한 문자열
    # dtype으로 재추론되어 이후 숫자 대입이 막히는 문제가 있어, object dtype으로 명시
    df["수량"] = pd.array(new_qty, dtype=object)
    return df


# ----------------------------------------------------------------------------
# 세션 상태 초기화
# ----------------------------------------------------------------------------
if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame(columns=COLUMNS)
if "df_backup" not in st.session_state:
    st.session_state.df_backup = None
if "df_backup_label" not in st.session_state:
    st.session_state.df_backup_label = ""


def backup_current_df(label):
    """데이터를 통째로 덮어쓰는 작업 전에 한 단계 백업을 저장해서
    '↩️ 되돌리기'로 복원할 수 있게 한다."""
    st.session_state.df_backup = st.session_state.df.copy()
    st.session_state.df_backup_label = label


st.title("🧪 시약 & 용액 재고 관리")
st.caption("PC와 모바일 브라우저에서 동일하게 열람·수정할 수 있습니다.")

# ----------------------------------------------------------------------------
# 0. Google Sheets 연동 (여러 사람이 같은 데이터를 공유해서 보고/수정)
#
# secrets.toml에 서비스 계정을 설정해두면 사이드바에서 시트를 불러오고
# 저장할 수 있습니다. 설정이 안 되어 있으면 이 부분은 비활성화되고,
# 엑셀 업로드/다운로드만으로도 그대로 사용 가능합니다.
# ----------------------------------------------------------------------------
gsheet_client = get_gsheet_client()

st.sidebar.header("🔗 Google Sheets 연동")

if gsheet_client is None:
    st.sidebar.info(
        "아직 Google Sheets 연동이 설정되지 않았습니다.\n\n"
        "여러 사람이 같은 재고 데이터를 함께 보고 수정하려면 "
        "README.md의 안내대로 서비스 계정을 설정해주세요.\n\n"
        "연동 전까지는 엑셀 업로드/다운로드로만 사용할 수 있습니다."
    )

try:
    default_sheet = st.secrets.get("gsheet", {}).get("spreadsheet_id", "")
except Exception:
    default_sheet = ""

sheet_id_input = st.sidebar.text_input(
    "스프레드시트 URL 또는 ID", value=default_sheet, key="sheet_id_input"
)
worksheet_name = st.sidebar.text_input(
    "워크시트(탭) 이름", value="재고", key="worksheet_name_input"
)

gsheet_ready = gsheet_client is not None and bool(sheet_id_input.strip())

# 최초 접속 시 시트 설정이 있으면 자동으로 한 번 불러옴
if gsheet_ready and "gsheet_autoloaded" not in st.session_state:
    try:
        ws = open_worksheet(gsheet_client, sheet_id_input, worksheet_name)
        st.session_state.df = recompute_quantity(load_df_from_worksheet(ws))
    except Exception as e:
        st.sidebar.error(f"자동 불러오기 실패: {e}")
    st.session_state.gsheet_autoloaded = True

sb_col1, sb_col2 = st.sidebar.columns(2)

if sb_col1.button("🔄 불러오기", use_container_width=True, disabled=not gsheet_ready):
    try:
        ws = open_worksheet(gsheet_client, sheet_id_input, worksheet_name)
        if not st.session_state.df.empty:
            backup_current_df("시트에서 불러오기 직전")
        st.session_state.df = recompute_quantity(load_df_from_worksheet(ws))
        st.sidebar.success(f"불러왔습니다 ({len(st.session_state.df)}행)")
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"불러오기 실패: {e}")

if sb_col2.button("💾 저장", use_container_width=True, disabled=not gsheet_ready):
    try:
        ws = open_worksheet(gsheet_client, sheet_id_input, worksheet_name)
        save_df_to_worksheet(ws, st.session_state.df)
        st.sidebar.success("시트에 저장했습니다!")
    except Exception as e:
        st.sidebar.error(f"저장 실패: {e}")

if gsheet_ready:
    st.sidebar.caption(
        "⚠️ '저장'을 눌러야 다른 사람에게도 내 수정사항이 반영됩니다.\n"
        "다른 사람이 저장한 최신 내용을 보려면 '불러오기'를 눌러주세요.\n"
        "(동시 저장 시 나중에 누른 저장이 우선 적용됩니다)"
    )

# ----------------------------------------------------------------------------
# 되돌리기(실행취소) UI
#
# "새로 열기"처럼 기존 데이터를 통째로 지우는 작업 직전에 자동으로 백업이
# 저장되어 있으면, 여기서 바로 복원할 수 있다.
# ----------------------------------------------------------------------------
undo_col1, undo_col2 = st.columns([1, 5])
with undo_col1:
    if st.button(
        "↩️ 되돌리기",
        disabled=st.session_state.df_backup is None,
        help="바로 직전 '새로 열기' 등 작업 이전 상태로 되돌립니다.",
    ):
        st.session_state.df, st.session_state.df_backup = (
            st.session_state.df_backup,
            st.session_state.df.copy(),
        )
        st.success("이전 상태로 되돌렸습니다.")
        st.rerun()
with undo_col2:
    if st.session_state.df_backup is not None:
        st.caption(f"↩️ 되돌릴 수 있는 백업 있음: {st.session_state.df_backup_label}")

# ----------------------------------------------------------------------------
# 1. 엑셀 불러오기 / 추가 불러오기(병합) / 저장
# ----------------------------------------------------------------------------
with st.expander("📁 엑셀 불러오기 / 병합 / 저장", expanded=True):
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("**새로 열기** (⚠️ 기존 데이터 전체 삭제)")
        new_file = st.file_uploader(
            "엑셀 파일 선택", type=["xlsx", "xls"], key="uploader_new"
        )
        confirm_overwrite = st.checkbox(
            "네, 기존 데이터를 지우고 새로 엽니다",
            key="confirm_overwrite",
            help="체크하지 않으면 '새로 열기' 버튼이 눌리지 않습니다. "
            "기존 데이터에 정보를 더하고 싶다면 대신 오른쪽 '추가 불러오기'를 쓰세요.",
        )
        if (
            new_file is not None
            and confirm_overwrite
            and st.button("📂 새로 열기 적용", key="btn_new")
        ):
            if not st.session_state.df.empty:
                backup_current_df("새로 열기 직전")
            df, missing = read_excel_normalized(new_file.getvalue(), COLUMNS)
            df = recompute_quantity(df)
            st.session_state.df = df
            msg = f"불러왔습니다! ({len(df)}행)"
            if missing:
                msg += f"\n\n※ 파일에 없어 빈 값으로 채운 컬럼: {', '.join(missing)}"
            msg += "\n\n혹시 실수하셨다면 상단 '↩️ 되돌리기'로 복원할 수 있습니다."
            st.success(msg)
            st.rerun()

    with col2:
        st.markdown("**추가 불러오기 (같은 항목은 최신 정보로 갱신)**")
        st.caption(
            "제품명+CAT No.가 같은 항목은 이 파일의 값으로 갱신되고, "
            "이 파일에 없는 항목(예: 직접 추가한 항목)은 그대로 유지됩니다. "
            "새 항목은 추가됩니다."
        )
        merge_file = st.file_uploader(
            "엑셀 파일 선택", type=["xlsx", "xls"], key="uploader_merge"
        )
        if merge_file is not None and st.button("➕ 병합 적용", key="btn_merge"):
            if not st.session_state.df.empty:
                backup_current_df("병합 적용 직전")
            new_df, missing = read_excel_normalized(merge_file.getvalue(), COLUMNS)
            if st.session_state.df.empty:
                merged_df = recompute_quantity(new_df)
                matched, updated_cells, added = 0, 0, len(new_df)
            else:
                merged_df, matched, updated_cells, added = merge_dataframes(
                    st.session_state.df, new_df, COLUMNS
                )
                merged_df = recompute_quantity(merged_df)
            st.session_state.df = merged_df
            msg = (
                f"병합 완료!\n- 매칭되어 최신 정보로 갱신된 항목: {matched}개 "
                f"(바뀐 칸 {updated_cells}개)\n- 새로 추가된 항목: {added}개\n"
                f"- 매칭 안 되는 기존 항목(수동 추가분 등)은 그대로 유지됨"
            )
            if missing:
                msg += f"\n\n※ 병합 파일에 없어 참고하지 않은 컬럼: {', '.join(missing)}"
            msg += "\n\n혹시 실수하셨다면 상단 '↩️ 되돌리기'로 복원할 수 있습니다."
            st.success(msg)
            st.rerun()

    with col3:
        st.markdown("**엑셀로 저장**")
        st.download_button(
            "💾 엑셀 다운로드",
            data=df_to_excel_bytes(st.session_state.df),
            file_name="lab_inventory.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        if st.button("🌐 빈 정보 자동 채우기 (PubChem API)", use_container_width=True):
            df = st.session_state.df.copy()
            today_str = datetime.date.today().strftime("%Y.%m.%d")
            updated, failed = 0, []
            headers = {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
                )
            }
            with st.spinner("PubChem에서 조회 중..."):
                for idx, row in df.iterrows():
                    if not str(row["등록일"]).strip():
                        df.at[idx, "등록일"] = today_str

                    cas_val = str(row["CAS No."]).strip()
                    name_val = str(row["제품명"]).strip()
                    query_id = cas_val if cas_val else name_val
                    need_hazard = not str(row["유해·위험성"]).strip()

                    if need_hazard and query_id:
                        try:
                            encoded_id = urllib.parse.quote(query_id)
                            url = (
                                "https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/"
                                f"name/{encoded_id}/property/Title,GHSClassification/JSON"
                            )
                            res = requests.get(url, headers=headers, timeout=5)
                            if res.status_code == 200:
                                data = res.json()
                                props = data["PropertyTable"]["Properties"][0]
                                ghs_list = props.get("GHSClassification", [])
                                hazard = ", ".join(ghs_list[:3]) if ghs_list else "GHS 정보 없음"
                                df.at[idx, "유해·위험성"] = hazard
                                updated += 1
                            else:
                                failed.append(f"{query_id} ({res.status_code})")
                        except Exception as e:
                            failed.append(f"{query_id} ({e})")
            st.session_state.df = df
            msg = f"자동 채우기 완료! 업데이트 {updated}건"
            if failed:
                msg += "\n\n실패 항목:\n" + "\n".join(failed[:5])
            st.success(msg)
            st.rerun()

# ----------------------------------------------------------------------------
# 2. 새 항목 추가 (용량 숫자+단위, 보관위치 선택, 미개봉/개봉 자동 수량 계산)
# ----------------------------------------------------------------------------
with st.expander("➕ 새 항목 추가", expanded=False):
    with st.form("add_item_form", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns(4)
        f_name = c1.text_input("제품명")
        f_maker = c2.text_input("제조사명")
        f_cat = c3.text_input("CAT No.")
        f_cas = c4.text_input("CAS No.")

        c5, c6, c7 = st.columns([1.2, 0.8, 1.5])
        f_cap_num = c5.text_input("용량 (숫자)")
        f_cap_unit = c6.selectbox("단위", CAPACITY_UNITS)
        f_storage = c7.selectbox("보관위치", STORAGE_LOCATIONS)

        c8, c9, c10, c11 = st.columns(4)
        f_unopened = c8.number_input("미개봉", min_value=0, step=1, value=0)
        f_opened = c9.number_input("개봉", min_value=0, step=1, value=0)
        c10.metric("수량 (자동 계산)", f_unopened + f_opened)
        f_reg_date = c11.text_input(
            "등록일", value=datetime.date.today().strftime("%Y.%m.%d")
        )

        c12, c13 = st.columns(2)
        f_expiry = c12.text_input("개봉 시 유통기한")
        f_hazard = c13.text_input("유해·위험성")

        submitted = st.form_submit_button("➕ 추가")
        if submitted:
            if not f_name.strip():
                st.warning("제품명을 입력해주세요.")
            else:
                capacity_text = (
                    f"{f_cap_num.strip()} {f_cap_unit}".strip() if f_cap_num.strip() else ""
                )
                new_row = {
                    "제품명": f_name.strip(),
                    "제조사명": f_maker.strip(),
                    "CAT No.": f_cat.strip(),
                    "CAS No.": f_cas.strip(),
                    "용량": capacity_text,
                    "수량": f_unopened + f_opened,
                    "미개봉": f_unopened,
                    "개봉": f_opened,
                    "보관위치": f_storage,
                    "개봉 시 유통기한": f_expiry.strip(),
                    "유해·위험성": f_hazard.strip(),
                    "등록일": f_reg_date.strip() or datetime.date.today().strftime("%Y.%m.%d"),
                }
                st.session_state.df = pd.concat(
                    [st.session_state.df, pd.DataFrame([new_row])], ignore_index=True
                )
                st.success("신규 항목이 추가되었습니다.")
                st.rerun()

# ----------------------------------------------------------------------------
# 3. 검색
# ----------------------------------------------------------------------------
search = st.text_input("🔍 검색어 (모든 컬럼에서 검색)")

if search.strip():
    mask = st.session_state.df.astype(str).apply(
        lambda col: col.str.contains(search, case=False, na=False)
    ).any(axis=1)
    view_df = st.session_state.df[mask]
    editable_mode = "fixed"  # 검색 중에는 행 추가/삭제 대신 값 수정만 허용
    st.caption(f"검색 결과 {len(view_df)}건 (행 추가/삭제는 검색어를 지운 뒤 이용해주세요)")
else:
    view_df = st.session_state.df
    editable_mode = "dynamic"

# ----------------------------------------------------------------------------
# 3.5 실험실 배치도 (표 바로 위) — 보관위치를 고르면 지도에서 강조됨
# ----------------------------------------------------------------------------
with st.expander("🗺️ 실험실 배치도 보기 (보관위치 선택 시 강조)", expanded=False):
    selected_location = st.selectbox(
        "보관위치",
        ["(전체 보기)"] + STORAGE_LOCATIONS,
        key="floor_plan_location_picker",
    )
    highlight = None if selected_location == "(전체 보기)" else selected_location
    components.html(build_floor_plan_svg(highlight), height=740, scrolling=True)

# ----------------------------------------------------------------------------
# 4. 데이터 표 (편집 가능) — 검색 없을 때는 표에서 바로 행 추가/삭제도 가능
# ----------------------------------------------------------------------------
edited_df = st.data_editor(
    view_df,
    num_rows=editable_mode,
    use_container_width=True,
    height=560,
    key="main_editor",
    column_config={
        "보관위치": st.column_config.SelectboxColumn(
            "보관위치", options=STORAGE_LOCATIONS + [""], required=False
        ),
        "수량": st.column_config.NumberColumn("수량"),
        "미개봉": st.column_config.NumberColumn("미개봉", min_value=0, step=1),
        "개봉": st.column_config.NumberColumn("개봉", min_value=0, step=1),
    },
)

if search.strip():
    # 검색(고정) 모드: 편집된 값만 원본에 반영, 행 추가/삭제는 없음
    for idx in edited_df.index:
        st.session_state.df.loc[idx] = edited_df.loc[idx]
    st.session_state.df = recompute_quantity(st.session_state.df)
else:
    st.session_state.df = recompute_quantity(edited_df.reset_index(drop=True))

st.caption(f"총 {len(st.session_state.df)}개 항목")
