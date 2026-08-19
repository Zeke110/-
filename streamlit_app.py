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
    """base_df에 new_df를 병합. 매칭되는 항목은 빈 칸만 채우고,
    매칭 안 되는 항목은 새 행으로 추가."""
    base_df = base_df.copy()

    key_to_idx = defaultdict(list)
    for idx, row in base_df.iterrows():
        key_to_idx[build_match_key(row)].append(idx)

    matched, filled = 0, 0
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
                    existing_empty = pd.isna(existing_val) or str(existing_val).strip() == ""
                    new_has_val = pd.notna(new_val) and str(new_val).strip() != ""
                    if existing_empty and new_has_val:
                        base_df.at[idx, col] = new_val
                        filled += 1
        else:
            new_rows.append(srow)

    if new_rows:
        base_df = pd.concat([base_df, pd.DataFrame(new_rows)], ignore_index=True)

    base_df = base_df.reset_index(drop=True)
    return base_df, matched, filled, len(new_rows)


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
# 1. 엑셀 불러오기 / 추가 불러오기(병합) / 저장
# ----------------------------------------------------------------------------
with st.expander("📁 엑셀 불러오기 / 병합 / 저장", expanded=True):
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("**새로 열기** (기존 데이터 덮어씀)")
        new_file = st.file_uploader(
            "엑셀 파일 선택", type=["xlsx", "xls"], key="uploader_new"
        )
        if new_file is not None and st.button("📂 새로 열기 적용", key="btn_new"):
            df, missing = read_excel_normalized(new_file.getvalue(), COLUMNS)
            df = recompute_quantity(df)
            st.session_state.df = df
            msg = f"불러왔습니다! ({len(df)}행)"
            if missing:
                msg += f"\n\n※ 파일에 없어 빈 값으로 채운 컬럼: {', '.join(missing)}"
            st.success(msg)
            st.rerun()

    with col2:
        st.markdown("**추가 불러오기 (병합)**")
        merge_file = st.file_uploader(
            "엑셀 파일 선택", type=["xlsx", "xls"], key="uploader_merge"
        )
        if merge_file is not None and st.button("➕ 병합 적용", key="btn_merge"):
            new_df, missing = read_excel_normalized(merge_file.getvalue(), COLUMNS)
            if st.session_state.df.empty:
                merged_df = recompute_quantity(new_df)
                matched, filled, added = 0, 0, len(new_df)
            else:
                merged_df, matched, filled, added = merge_dataframes(
                    st.session_state.df, new_df, COLUMNS
                )
                merged_df = recompute_quantity(merged_df)
            st.session_state.df = merged_df
            msg = (
                f"병합 완료!\n- 매칭되어 정보가 보완된 항목: {matched}개\n"
                f"- 채워진 빈 칸 수: {filled}개\n- 새로 추가된 항목: {added}개"
            )
            if missing:
                msg += f"\n\n※ 병합 파일에 없어 참고하지 않은 컬럼: {', '.join(missing)}"
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
        "수량": st.column_config.NumberColumn(
            "수량 (미개봉+개봉 자동계산, 필요시 직접 수정 가능)"
        ),
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
