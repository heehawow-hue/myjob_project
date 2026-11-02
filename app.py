"""스탠다드 뱅크 인증 프로젝트 접수·진행현황 대시보드 (교육용 가상 데이터)"""

from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ───────────────────────── 기본 설정 ─────────────────────────
st.set_page_config(page_title="인증 프로젝트 진행현황", page_icon="📋", layout="wide")

DATA_DIR = Path(__file__).parent / "data"
CANDIDATE_FILES = ["SB_projects_demo.csv", "SB_project_demo.csv"]
DEFAULT_BASE_DATE = date(2026, 10, 23)

COLUMNS = ["의뢰번호", "고객사", "제품명", "모델명", "시험분야", "적용규격",
           "담당엔지니어", "접수일", "완료예정일", "진행상태", "시료상태", "메모"]
STATUSES = ["접수", "자료확인", "시험진행", "성적서검토", "보완요청", "완료"]
FILTER_FIELDS = ["시험분야", "적용규격", "담당엔지니어", "진행상태", "시료상태"]
DEFAULT_TEST_FIELDS = ["EMC", "무선(RF)", "인증서류", "SAR"]

# 일정 상태 라벨
SCHEDULE_DONE = "완료"
SCHEDULE_LATE = "지연 후보(일정 확인 필요)"
SCHEDULE_TODAY = "오늘 마감"
SCHEDULE_PLANNED = "진행 예정"
SCHEDULE_UNKNOWN = "일정 확인 필요"
SCHEDULE_LABELS = [SCHEDULE_LATE, SCHEDULE_TODAY, SCHEDULE_PLANNED, SCHEDULE_UNKNOWN, SCHEDULE_DONE]

EXCEL_NOTICE = "교육용 검토 파일 — 공식 보고서 아님"
SESSION_NOTICE = ("시험분야 등록·변경·삭제 및 재분류 내용은 현재 브라우저 세션(session_state)에만 "
                  "보관됩니다. 새로고침하거나 세션이 종료되면 사라질 수 있으며, 원본 CSV는 변경되지 않습니다.")

NAVY = "#0B2545"
BLUE = "#1F5FBF"
LIGHT = "#EEF3FB"
LATE_BG = "#FDE2E1"
LATE_TEXT = "#9B1C1C"
# 진행상태별 고정 색 (순서 고정 — 필터로 일부 상태가 빠져도 색이 바뀌지 않음)
STATUS_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300"]

# ───────────────────────── 스타일 ─────────────────────────
st.markdown(
    f"""
    <style>
    .stApp {{ background-color: #FFFFFF; }}
    h1, h2, h3 {{ color: {NAVY}; }}
    .demo-banner {{
        background: {LIGHT}; border-left: 6px solid {BLUE}; color: {NAVY};
        padding: 10px 16px; border-radius: 6px; font-weight: 600; margin-bottom: 12px;
    }}
    .card {{
        background: #FFFFFF; border: 1px solid #D6E0F0; border-top: 4px solid {BLUE};
        border-radius: 10px; padding: 14px 10px; text-align: center;
        box-shadow: 0 1px 3px rgba(11,37,69,0.08);
    }}
    .card.total {{ background: {NAVY}; border-top-color: {NAVY}; }}
    .card .label {{ font-size: 0.9rem; color: #4A5A78; }}
    .card .value {{ font-size: 1.9rem; font-weight: 700; color: {NAVY}; }}
    .card.total .label, .card.total .value {{ color: #FFFFFF; }}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="demo-banner">⚠️ 교육용 가상 데이터 기준 — 실제 업무 시스템 아님</div>',
    unsafe_allow_html=True,
)
st.title("📋 인증 프로젝트 접수·진행현황 대시보드")


# ───────────────────────── 데이터 로드 ─────────────────────────
def find_data_file():
    for name in CANDIDATE_FILES:
        path = DATA_DIR / name
        if path.exists():
            return path
    return None


@st.cache_data
def load_csv(source):
    df = pd.read_csv(source, encoding="utf-8-sig", dtype=str)
    df.columns = [c.strip() for c in df.columns]
    for col in ["접수일", "완료예정일"]:
        if col in df.columns:
            # 비어 있거나 날짜 형식이 아니면 NaT
            df[col] = pd.to_datetime(df[col], errors="coerce", format="mixed").dt.normalize()
    return df


data_path = find_data_file()
raw = None

if data_path is not None:
    raw = load_csv(data_path)
    source_id = f"data/{data_path.name}"
else:
    st.info(
        "data 폴더에서 SB_projects_demo.csv 파일을 찾지 못했습니다. "
        "아래에서 CSV 파일을 직접 업로드해 주세요."
    )
    uploaded = st.file_uploader("의뢰 목록 CSV 업로드", type=["csv"])
    if uploaded is not None:
        raw = load_csv(uploaded)
        source_id = f"업로드: {uploaded.name}"

if raw is None:
    st.stop()

missing = [c for c in COLUMNS if c not in raw.columns]
if missing:
    st.warning(f"CSV에 다음 필드가 없습니다: {', '.join(missing)}")
    st.stop()

# ───────────────────────── 세션 데이터 (카테고리 관리 대상) ─────────────────────────
# 데이터 출처가 바뀔 때만 원본으로 초기화하고, 이후 변경은 session_state에서 유지
if st.session_state.get("source_id") != source_id:
    st.session_state.source_id = source_id
    st.session_state.projects = raw.copy()
    extra = sorted(f for f in raw["시험분야"].dropna().unique() if f not in DEFAULT_TEST_FIELDS)
    st.session_state.test_fields = DEFAULT_TEST_FIELDS + extra

projects = st.session_state.projects
test_fields = st.session_state.test_fields

st.caption(f"데이터 출처: {source_id}")
st.caption(f"💾 {SESSION_NOTICE}")


def flash(level, text):
    st.session_state.cat_msg = (level, text)


def replace_in_filter(old, new=None):
    key = "flt_시험분야"
    picked = st.session_state.get(key, [])
    if old in picked:
        st.session_state[key] = [new if v == old else v for v in picked if new or v != old]


def add_field():
    name = st.session_state.new_field.strip()
    if not name:
        flash("warning", "등록할 시험분야 이름을 입력해 주세요.")
    elif name in test_fields:
        flash("warning", f"'{name}'은(는) 이미 등록된 시험분야입니다.")
    else:
        test_fields.append(name)
        st.session_state.new_field = ""
        flash("success", f"'{name}' 시험분야를 등록했습니다.")


def rename_field():
    old = st.session_state.rename_src
    new = st.session_state.rename_new.strip()
    if not new:
        flash("warning", "새 이름을 입력해 주세요.")
    elif new == old:
        flash("info", "기존 이름과 같습니다.")
    elif new in test_fields:
        flash("warning", f"'{new}'은(는) 이미 등록된 시험분야입니다. 다른 이름을 입력해 주세요.")
    else:
        test_fields[test_fields.index(old)] = new
        mask = projects["시험분야"] == old
        projects.loc[mask, "시험분야"] = new
        replace_in_filter(old, new)
        st.session_state.rename_src = new
        st.session_state.rename_new = ""
        flash("success", f"'{old}' → '{new}' 이름 변경 완료 (관련 의뢰 {int(mask.sum())}건 함께 갱신)")


def reclassify_field():
    src = st.session_state.delete_src
    dst = st.session_state.reassign_dst
    mask = projects["시험분야"] == src
    projects.loc[mask, "시험분야"] = dst
    flash("success", f"'{src}' 의뢰 {int(mask.sum())}건을 '{dst}'(으)로 재분류했습니다. 이제 삭제할 수 있습니다.")


def delete_field():
    name = st.session_state.delete_src
    if (projects["시험분야"] == name).any():
        flash("error", f"'{name}'은(는) 사용 중이라 삭제할 수 없습니다. 먼저 다른 분야로 재분류해 주세요.")
        return
    test_fields.remove(name)
    replace_in_filter(name)
    st.session_state.pop("delete_src", None)
    flash("success", f"'{name}' 시험분야를 삭제했습니다.")


# ───────────────────────── 사이드바: ① 필터 ─────────────────────────
with st.sidebar:
    st.header("🔎 필터")
    base_date = st.date_input("기준일", value=DEFAULT_BASE_DATE)
    only_received = st.checkbox("기준일까지 접수된 건만 보기", value=True)

    selections = {}
    for field in FILTER_FIELDS:
        if field == "시험분야":
            options = list(test_fields)
        elif field == "진행상태":
            options = [s for s in STATUSES if s in set(projects[field].dropna())]
        else:
            options = sorted(projects[field].dropna().unique())
        key = f"flt_{field}"
        # 카테고리 변경으로 사라진 값은 선택에서 제거
        st.session_state[key] = [v for v in st.session_state.get(key, []) if v in options]
        selections[field] = st.multiselect(field, options, key=key, placeholder="전체")

# ───────────────────────── 사이드바: ② 카테고리 관리 ─────────────────────────
with st.sidebar:
    st.divider()
    st.header("🗂️ 카테고리 관리")
    st.caption("시험분야 등록 / 이름 변경 / 삭제")

    usage = projects["시험분야"].value_counts()
    st.markdown("\n".join(f"- {f} — {int(usage.get(f, 0))}건" for f in test_fields) or "_등록된 시험분야 없음_")

    action = st.radio("작업", ["등록", "이름 변경", "삭제"], horizontal=True, key="cat_action")

    if action == "등록":
        st.text_input("새 시험분야 이름", key="new_field", placeholder="예: 전기안전")
        st.button("등록", on_click=add_field, type="primary", width="stretch")

    elif not test_fields:
        st.info("등록된 시험분야가 없습니다.")

    elif action == "이름 변경":
        st.selectbox("변경할 시험분야", test_fields, key="rename_src")
        st.text_input("새 이름", key="rename_new")
        st.button("이름 변경", on_click=rename_field, type="primary", width="stretch")
        st.caption("이름을 바꾸면 해당 분야의 의뢰 데이터도 함께 갱신됩니다.")

    else:
        target = st.selectbox("삭제할 시험분야", test_fields, key="delete_src")
        used = int(usage.get(target, 0))
        if used:
            st.warning(
                f"'{target}'은(는) 의뢰 {used}건에서 사용 중이라 바로 삭제할 수 없습니다. "
                "아래에서 다른 분야로 재분류한 뒤 삭제해 주세요."
            )
            others = [f for f in test_fields if f != target]
            if others:
                st.selectbox("재분류할 시험분야", others, key="reassign_dst")
                st.button(f"{used}건 재분류", on_click=reclassify_field, width="stretch")
            else:
                st.info("재분류할 다른 시험분야가 없습니다. 먼저 새 시험분야를 등록해 주세요.")
        st.button("삭제", on_click=delete_field, disabled=bool(used), type="primary", width="stretch")

    msg = st.session_state.pop("cat_msg", None)
    if msg:
        getattr(st, msg[0])(msg[1])

    st.caption(f"⚠️ {SESSION_NOTICE}")


# ───────────────────────── 잔여일수·일정 상태 ─────────────────────────
def schedule_status(row):
    if row["진행상태"] == "완료":
        return SCHEDULE_DONE
    remaining = row["잔여일수"]
    if pd.isna(remaining):
        return SCHEDULE_UNKNOWN
    if remaining < 0:
        return SCHEDULE_LATE
    if remaining == 0:
        return SCHEDULE_TODAY
    return SCHEDULE_PLANNED


def with_schedule(df, base):
    df = df.copy()
    df["잔여일수"] = (df["완료예정일"] - pd.Timestamp(base)).dt.days.astype("Int64")
    df["일정 상태"] = df.apply(schedule_status, axis=1) if len(df) else pd.Series(dtype=str)
    return df


all_rows = with_schedule(projects, base_date)

view = all_rows
if only_received:
    view = view[view["접수일"].notna() & (view["접수일"] <= pd.Timestamp(base_date))]
for field, picked in selections.items():
    if picked:
        view = view[view[field].isin(picked)]

# ───────────────────────── 요약 카드 ─────────────────────────
st.subheader(f"요약 (기준일 {base_date:%Y-%m-%d})")


def card(col, label, value, total=False):
    cls = "card total" if total else "card"
    col.markdown(
        f'<div class="{cls}"><div class="label">{label}</div>'
        f'<div class="value">{value}</div></div>',
        unsafe_allow_html=True,
    )


counts = view["진행상태"].value_counts()
cols = st.columns(len(STATUSES) + 1)
card(cols[0], "전체", f"{len(view)}건", total=True)
for col, status in zip(cols[1:], STATUSES):
    card(col, status, f"{int(counts.get(status, 0))}건")

if view.empty:
    st.info("조건에 맞는 인증 프로젝트가 없습니다")
else:
    late = int((view["일정 상태"] == SCHEDULE_LATE).sum())
    unknown = int((view["일정 상태"] == SCHEDULE_UNKNOWN).sum())
    if late:
        st.error(f"🔴 지연 후보(일정 확인 필요): {late}건 — 완료예정일이 기준일보다 지난 미완료 건")
    if unknown:
        st.warning(f"⚠️ 완료예정일이 비어 있거나 형식이 잘못된 미완료 건: {unknown}건")

    # ───────────────────────── 차트 ─────────────────────────
    left, right = st.columns(2)

    with left:
        st.subheader("시험분야별 건수")
        by_field = view["시험분야"].value_counts().rename_axis("시험분야").reset_index(name="건수")
        bar = (
            alt.Chart(by_field)
            .mark_bar(color=BLUE, cornerRadiusEnd=4, size=26)
            .encode(
                y=alt.Y("시험분야:N", sort="-x", title=None),
                x=alt.X("건수:Q", title="건수", axis=alt.Axis(tickMinStep=1, grid=True)),
                tooltip=["시험분야:N", "건수:Q"],
            )
            .properties(height=max(160, 46 * len(by_field)))
            .configure_axis(labelColor=NAVY, titleColor="#4A5A78", gridColor="#E6ECF5",
                            domainColor="#C9D3E3")
            .configure_view(stroke=None)
        )
        st.altair_chart(bar, width="stretch")

    with right:
        st.subheader("진행상태 비율")
        by_status = view["진행상태"].value_counts().rename_axis("진행상태").reset_index(name="건수")
        by_status["비율"] = by_status["건수"] / by_status["건수"].sum()
        donut = (
            alt.Chart(by_status)
            .mark_arc(innerRadius=60, outerRadius=110, stroke="#FFFFFF", strokeWidth=2)
            .encode(
                theta=alt.Theta("건수:Q", stack=True),
                color=alt.Color(
                    "진행상태:N",
                    scale=alt.Scale(domain=STATUSES, range=STATUS_COLORS),
                    legend=alt.Legend(title=None, orient="right", labelColor=NAVY),
                ),
                order=alt.Order("진행상태_순서:Q"),
                tooltip=["진행상태:N", "건수:Q", alt.Tooltip("비율:Q", format=".0%")],
            )
            .transform_calculate(진행상태_순서=f"indexof({STATUSES}, datum.진행상태)")
            .properties(height=260)
            .configure_view(stroke=None)
        )
        st.altair_chart(donut, width="stretch")

    # ───────────────────────── 의뢰 목록 ─────────────────────────
    st.subheader("의뢰 목록")

    table = view[COLUMNS + ["잔여일수", "일정 상태"]].sort_values("접수일")

    def highlight_late(row):
        style = (f"background-color: {LATE_BG}; color: {LATE_TEXT}"
                 if row["일정 상태"] == SCHEDULE_LATE else "")
        return [style] * len(row)

    st.dataframe(
        table.style.apply(highlight_late, axis=1),
        hide_index=True,
        width="stretch",
        column_config={
            "접수일": st.column_config.DateColumn("접수일", format="YYYY-MM-DD"),
            "완료예정일": st.column_config.DateColumn("완료예정일", format="YYYY-MM-DD"),
            "잔여일수": st.column_config.NumberColumn(
                "잔여일수", help="완료예정일 - 기준일 (음수면 예정일 경과)"),
            "일정 상태": st.column_config.TextColumn("일정 상태", width="medium"),
        },
    )
    st.caption(f"표시 {len(table)}건 / 전체 {len(projects)}건 · 빨간색 행 = 지연 후보(일정 확인 필요)")


# ───────────────────────── Excel 다운로드 ─────────────────────────
def project_sheet(df):
    out = df[COLUMNS + ["잔여일수", "일정 상태"]].sort_values("접수일").copy()
    for col in ["접수일", "완료예정일"]:
        out[col] = out[col].dt.strftime("%Y-%m-%d")
    out["잔여일수"] = out["잔여일수"].astype(object).where(out["잔여일수"].notna(), None)
    return out


def status_sheet(df):
    by_status = df["진행상태"].value_counts().reindex(STATUSES, fill_value=0)
    status = by_status.rename_axis("진행상태").reset_index(name="건수")
    by_schedule = df["일정 상태"].value_counts().reindex(SCHEDULE_LABELS, fill_value=0)
    schedule = by_schedule.rename_axis("일정 상태").reset_index(name="건수")
    total = max(len(df), 1)
    for t in (status, schedule):
        t["비율"] = t["건수"] / total
    status.loc[len(status)] = ["합계", len(df), 1.0 if len(df) else 0.0]
    return [status, schedule]


def engineer_sheet(df):
    pivot = pd.crosstab(df["담당엔지니어"], df["진행상태"]).reindex(columns=STATUSES, fill_value=0)
    pivot["합계"] = pivot.sum(axis=1)
    pivot["지연 후보"] = (
        df[df["일정 상태"] == SCHEDULE_LATE].groupby("담당엔지니어").size()
        .reindex(pivot.index, fill_value=0)
    )
    return pivot.reset_index().rename(columns={"index": "담당엔지니어"})


def condition_sheet(scope, df, filtered):
    rows = [
        ("구분", scope),
        ("기준일", f"{base_date:%Y-%m-%d}"),
        ("다운로드 일시", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("데이터 출처", source_id),
        ("기준일까지 접수된 건만", ("적용" if only_received else "미적용") if filtered else "미적용 (전체 목록)"),
    ]
    for field in FILTER_FIELDS:
        picked = selections[field] if filtered else []
        rows.append((f"필터: {field}", ", ".join(picked) if picked else "전체"))
    rows += [
        ("건수", f"{len(df)}건"),
        ("등록된 시험분야", ", ".join(test_fields)),
        ("비고", "교육용 가상 데이터. 세션 중 변경한 시험분야(이름 변경·재분류)가 반영된 결과입니다."),
    ]
    return [pd.DataFrame(rows, columns=["항목", "값"])]


def text_width(value):
    s = "" if value is None else str(value)
    return sum(2 if ord(ch) > 127 else 1 for ch in s)


def write_sheet(writer, name, tables):
    row = 2  # 0-based: 1행 안내 문구, 2행 공백, 3행부터 표
    header_rows = []
    for t in tables:
        t.to_excel(writer, sheet_name=name, index=False, startrow=row)
        header_rows.append(row + 1)
        row += len(t) + 2
    ws = writer.sheets[name]

    ws["A1"] = EXCEL_NOTICE
    ws["A1"].font = Font(bold=True, color="C00000", size=12)

    header_fill = PatternFill("solid", fgColor=NAVY.lstrip("#"))
    for r in header_rows:
        for cell in ws[r]:
            if cell.value is not None:
                cell.fill = header_fill
                cell.font = Font(bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center")

    for col_cells in ws.iter_cols(min_row=3):
        width = max(text_width(c.value) for c in col_cells)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max(width + 2, 8), 60)
        for c in col_cells:
            if isinstance(c.value, float) and c.column_letter == "C" and name == "상태별집계":
                c.number_format = "0.0%"
    ws.freeze_panes = "A4"


def build_excel(df, scope, filtered):
    sheets = {
        "프로젝트목록": [project_sheet(df)],
        "상태별집계": status_sheet(df),
        "담당자별집계": [engineer_sheet(df)],
        "분석조건": condition_sheet(scope, df, filtered),
    }
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, tables in sheets.items():
            write_sheet(writer, name, tables)
        # 프로젝트목록의 지연 후보 행 강조
        ws = writer.sheets["프로젝트목록"]
        status_col = len(COLUMNS) + 2
        late_fill = PatternFill("solid", fgColor=LATE_BG.lstrip("#"))
        for r in range(4, ws.max_row + 1):
            if ws.cell(r, status_col).value == SCHEDULE_LATE:
                for c in ws[r]:
                    c.fill = late_fill
                    c.font = Font(color=LATE_TEXT.lstrip("#"))
    return buf.getvalue()


st.divider()
st.subheader("📥 Excel 다운로드")
stamp = date.today().strftime("%Y%m%d")  # 실제 다운로드(생성) 날짜
st.caption(f"시트: 프로젝트목록 / 상태별집계 / 담당자별집계 / 분석조건 · 각 시트 상단에 '{EXCEL_NOTICE}' 표시")

d1, d2 = st.columns(2)
d1.download_button(
    f"전체 목록 ({len(all_rows)}건)",
    data=build_excel(all_rows, "전체 목록", filtered=False),
    file_name=f"standardbank_project_status_{stamp}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    width="stretch",
)
d2.download_button(
    f"현재 필터 결과 ({len(view)}건)",
    data=build_excel(view, "현재 필터 결과", filtered=True) if len(view) else b"",
    file_name=f"standardbank_project_status_filtered_{stamp}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    disabled=view.empty,
    help="필터 결과가 0건이면 다운로드할 수 없습니다." if view.empty else None,
    type="primary",
    width="stretch",
)
