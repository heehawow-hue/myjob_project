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
PROJECT_FILES = ["SB_projects_demo.csv", "SB_project_demo.csv"]
RESOURCE_FILES = ["SB_resources_demo.csv", "SB_resource_demo.csv"]
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

DEMO_NOTICE = "⚠️ 교육용 가상 데이터 기준 — 실제 업무 시스템 아님"
EXCEL_NOTICE = "교육용 검토 파일 — 공식 보고서 아님"
RESOURCE_NOTICE = "현황 조회용이며 예약·시험 재개를 승인하는 기능이 아님"
SESSION_NOTICE = ("시험분야 등록·변경·삭제, 재분류, 메모 수정 내용은 현재 브라우저 세션(session_state)에만 "
                  "보관됩니다. 새로고침하거나 세션이 종료되면 사라질 수 있으며, 원본 CSV는 변경되지 않습니다.")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

NAVY = "#0B2545"
BLUE = "#1F5FBF"
LIGHT = "#EEF3FB"
LATE_BG = "#FDE2E1"
LATE_TEXT = "#9B1C1C"
# 진행상태별 고정 색 (순서 고정 — 필터로 일부 상태가 빠져도 색이 바뀌지 않음)
STATUS_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300"]
# 자원 상태별 강조 색 (상태 글자와 함께 표시하므로 색만으로 구분하지 않음)
RESOURCE_TONES = {"사용 가능": "good", "사용 중": "busy", "점검 중": "alert"}

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
        box-shadow: 0 1px 3px rgba(11,37,69,0.08); height: 100%;
    }}
    .card.total {{ background: {NAVY}; border-top-color: {NAVY}; }}
    .card .label {{ font-size: 0.9rem; color: #4A5A78; }}
    .card .value {{ font-size: 1.9rem; font-weight: 700; color: {NAVY}; }}
    .card .sub {{ font-size: 0.8rem; color: #4A5A78; margin-top: 4px; }}
    .card.total .label, .card.total .value {{ color: #FFFFFF; }}
    .card.good {{ border-top-color: #1E7B34; }}  .card.good .value {{ color: #1E7B34; }}
    .card.busy {{ border-top-color: {BLUE}; }}   .card.busy .value {{ color: {BLUE}; }}
    .card.alert {{ border-top-color: #C0392B; }} .card.alert .value {{ color: #C0392B; }}
    .card.unknown {{ border-top-color: #8A94A6; }} .card.unknown .value {{ color: #4A5A78; }}
    .detail {{ background: {LIGHT}; border-radius: 10px; padding: 14px 18px; }}
    .detail dt {{ color: #4A5A78; font-size: 0.85rem; margin-top: 6px; }}
    .detail dd {{ color: {NAVY}; font-weight: 600; margin: 0 0 4px 0; }}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(f'<div class="demo-banner">{DEMO_NOTICE}</div>', unsafe_allow_html=True)
st.title("📋 인증 프로젝트 접수·진행현황 대시보드")


# ───────────────────────── 공통 함수 ─────────────────────────
def find_data_file(candidates):
    """data 폴더에서 후보 파일명 중 처음 존재하는 파일 경로를 반환 (없으면 None)."""
    for name in candidates:
        path = DATA_DIR / name
        if path.exists():
            return path
    return None


@st.cache_data
def load_csv(source, date_cols=()):
    """CSV를 문자열로 읽고 지정한 날짜 열을 날짜형으로 변환 (형식 오류·빈 값은 NaT)."""
    df = pd.read_csv(source, encoding="utf-8-sig", dtype=str)
    df.columns = [c.strip() for c in df.columns]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", format="mixed").dt.normalize()
    return df


def card(col, label, value, sub="", tone=""):
    """요약·자원 카드 한 장을 HTML로 그림. tone은 total/good/busy/alert/unknown 중 하나."""
    sub_html = f'<div class="sub">{sub}</div>' if sub else ""
    col.markdown(
        f'<div class="card {tone}"><div class="label">{label}</div>'
        f'<div class="value">{value}</div>{sub_html}</div>',
        unsafe_allow_html=True,
    )


def flash(level, text):
    """콜백에서 만든 안내 메시지를 다음 화면 갱신 때 한 번 표시하도록 보관."""
    st.session_state.cat_msg = (level, text)


def keep_valid_selection(key, options):
    """다중 선택 위젯 상태에서 현재 선택지에 없는 값(삭제·변경된 값)을 제거."""
    st.session_state[key] = [v for v in st.session_state.get(key, []) if v in options]


def fmt_date(value):
    """날짜를 YYYY-MM-DD 문자열로 표시 (비어 있으면 안내 문구)."""
    return value.strftime("%Y-%m-%d") if pd.notna(value) else "미정(일정 확인 필요)"


# ───────────────────────── 데이터 로드 ─────────────────────────
data_path = find_data_file(PROJECT_FILES)
raw = None

if data_path is not None:
    raw = load_csv(data_path, ("접수일", "완료예정일"))
    source_id = f"data/{data_path.name}"
else:
    st.info(
        "data 폴더에서 SB_projects_demo.csv 파일을 찾지 못했습니다. "
        "아래에서 CSV 파일을 직접 업로드해 주세요."
    )
    uploaded = st.file_uploader("의뢰 목록 CSV 업로드", type=["csv"])
    if uploaded is not None:
        raw = load_csv(uploaded, ("접수일", "완료예정일"))
        source_id = f"업로드: {uploaded.name}"

if raw is None:
    st.stop()

missing = [c for c in COLUMNS if c not in raw.columns]
if missing:
    st.warning(f"CSV에 다음 필드가 없습니다: {', '.join(missing)}")
    st.stop()

# ───────────────────────── 세션 데이터 (카테고리·메모 변경 대상) ─────────────────────────
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


# ───────────────────────── 카테고리 관리 콜백 ─────────────────────────
def replace_in_filter(old, new=None):
    """시험분야 필터 선택값에서 old를 new로 바꾸거나(new=None이면) 제거."""
    key = "flt_시험분야"
    picked = st.session_state.get(key, [])
    if old in picked:
        st.session_state[key] = [new if v == old else v for v in picked if new or v != old]


def add_field():
    """시험분야 등록: 빈 이름·중복 이름은 막고 목록 끝에 추가."""
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
    """시험분야 이름 변경: 목록·관련 의뢰 데이터·필터 선택값을 함께 갱신."""
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
    """삭제하려는 시험분야의 의뢰를 선택한 다른 시험분야로 일괄 재분류."""
    src = st.session_state.delete_src
    dst = st.session_state.reassign_dst
    mask = projects["시험분야"] == src
    projects.loc[mask, "시험분야"] = dst
    flash("success", f"'{src}' 의뢰 {int(mask.sum())}건을 '{dst}'(으)로 재분류했습니다. 이제 삭제할 수 있습니다.")


def delete_field():
    """시험분야 삭제: 사용 중인 분야는 삭제하지 않고 재분류를 안내."""
    name = st.session_state.delete_src
    if (projects["시험분야"] == name).any():
        flash("error", f"'{name}'은(는) 사용 중이라 삭제할 수 없습니다. 먼저 다른 분야로 재분류해 주세요.")
        return
    test_fields.remove(name)
    replace_in_filter(name)
    st.session_state.pop("delete_src", None)
    flash("success", f"'{name}' 시험분야를 삭제했습니다.")


def save_memo(request_id):
    """상세 영역의 메모 입력값을 해당 의뢰의 메모로 저장 (session_state 데이터만 변경)."""
    text = st.session_state[f"memo_{request_id}"].strip()
    projects.loc[projects["의뢰번호"] == request_id, "메모"] = text
    st.session_state.memo_msg = f"{request_id} 메모를 저장했습니다. (현재 세션에만 보관)"


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
        keep_valid_selection(key, options)
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


# ───────────────────────── 시험 자원 현황 ─────────────────────────
def render_resources():
    """자원 CSV의 상태를 그대로 카드로 표시. '점검 중' 등 상태 값은 코드에서 바꾸지 않음."""
    st.subheader("🧪 시험 자원 현황")
    path = find_data_file(RESOURCE_FILES)
    if path is None:
        st.info("data 폴더에서 SB_resources_demo.csv 파일을 찾지 못해 자원 현황을 표시하지 않습니다.")
        return
    resources = load_csv(path)
    if not {"자원", "상태"} <= set(resources.columns):
        st.warning("자원 CSV에 '자원', '상태' 필드가 필요합니다.")
        return

    cols = st.columns(max(len(resources), 1))
    for col, (_, row) in zip(cols, resources.iterrows()):
        status = row["상태"] if pd.notna(row["상태"]) else "상태 미기재"
        note = row.get("비고")
        sub = (f"{note}<br>" if pd.notna(note) else "") + RESOURCE_NOTICE
        card(col, row["자원"], status, sub=sub, tone=RESOURCE_TONES.get(status, "unknown"))
    st.caption(f"ℹ️ {RESOURCE_NOTICE}. 데이터 파일: data/{path.name}")


render_resources()


# ───────────────────────── 잔여일수·일정 상태 ─────────────────────────
def schedule_status(row):
    """한 의뢰의 일정 상태 판정: 완료 → 일정 없음 → 지연 후보 → 오늘 마감 → 진행 예정 순."""
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
    """잔여일수(완료예정일 - 기준일)와 일정 상태 열을 붙인 사본을 반환."""
    df = df.copy()
    df["잔여일수"] = (df["완료예정일"] - pd.Timestamp(base)).dt.days.astype("Int64")
    df["일정 상태"] = df.apply(schedule_status, axis=1) if len(df) else pd.Series(dtype=str)
    return df


def apply_filters(df):
    """기준일 접수 조건과 사이드바 다중 선택 필터를 적용한 결과를 반환."""
    if only_received:
        df = df[df["접수일"].notna() & (df["접수일"] <= pd.Timestamp(base_date))]
    for field, picked in selections.items():
        if picked:
            df = df[df[field].isin(picked)]
    return df


all_rows = with_schedule(projects, base_date)
view = apply_filters(all_rows)

# ───────────────────────── 요약 카드 ─────────────────────────
st.subheader(f"📊 진행 요약 (기준일 {base_date:%Y-%m-%d})")

counts = view["진행상태"].value_counts()
cols = st.columns(len(STATUSES) + 1)
card(cols[0], "전체", f"{len(view)}건", tone="total")
for col, status in zip(cols[1:], STATUSES):
    card(col, status, f"{int(counts.get(status, 0))}건")


# ───────────────────────── 차트 ─────────────────────────
def field_bar_chart(df):
    """시험분야별 건수 가로 막대그래프 (건수 내림차순)."""
    by_field = df["시험분야"].value_counts().rename_axis("시험분야").reset_index(name="건수")
    return (
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


def status_donut_chart(df):
    """진행상태 비율 도넛 차트 (상태별 색 고정, 범례·툴팁 포함)."""
    by_status = df["진행상태"].value_counts().rename_axis("진행상태").reset_index(name="건수")
    by_status["비율"] = by_status["건수"] / by_status["건수"].sum()
    return (
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


def highlight_late(row):
    """표에서 지연 후보 행을 빨간색 계열로 강조하는 행 스타일."""
    style = (f"background-color: {LATE_BG}; color: {LATE_TEXT}"
             if row["일정 상태"] == SCHEDULE_LATE else "")
    return [style] * len(row)


def render_detail(df):
    """의뢰번호를 선택하면 상세 정보를 보여 주고, 메모를 입력·저장할 수 있게 함."""
    st.subheader("🔍 의뢰 상세조회")
    ids = df.sort_values("접수일")["의뢰번호"].tolist()
    if st.session_state.get("detail_id") not in ids:
        st.session_state.pop("detail_id", None)
    request_id = st.selectbox("의뢰번호 선택 (현재 필터 결과 기준)", ids, key="detail_id")
    row = df[df["의뢰번호"] == request_id].iloc[0]

    info, memo = st.columns([3, 2])
    items = [
        ("고객사", row["고객사"]),
        ("제품 · 모델", f"{row['제품명']} / {row['모델명']}"),
        ("시험분야", row["시험분야"]),
        ("적용규격", row["적용규격"]),
        ("담당엔지니어", row["담당엔지니어"]),
        ("시료상태", row["시료상태"]),
        ("완료예정일", f"{fmt_date(row['완료예정일'])} · {row['일정 상태']}"),
    ]
    half = (len(items) + 1) // 2
    left_html = "".join(f"<dt>{k}</dt><dd>{v}</dd>" for k, v in items[:half])
    right_html = "".join(f"<dt>{k}</dt><dd>{v}</dd>" for k, v in items[half:])
    info.markdown(
        f'<div class="detail" style="display:flex;gap:32px">'
        f"<dl style='flex:1;margin:0'>{left_html}</dl><dl style='flex:1;margin:0'>{right_html}</dl></div>",
        unsafe_allow_html=True,
    )

    current = row["메모"] if pd.notna(row["메모"]) else ""
    with memo.form(f"memo_form_{request_id}", border=False):
        st.text_area("담당자 메모", value=current, key=f"memo_{request_id}", height=140)
        st.form_submit_button("메모 저장", on_click=save_memo, args=(request_id,), type="primary")
    saved = st.session_state.pop("memo_msg", None)
    if saved:
        memo.success(saved)
    memo.caption("메모 수정 내용은 현재 세션에만 보관되며 원본 CSV에는 저장되지 않습니다.")


if view.empty:
    st.info("조건에 맞는 인증 프로젝트가 없습니다")
else:
    late = int((view["일정 상태"] == SCHEDULE_LATE).sum())
    unknown = int((view["일정 상태"] == SCHEDULE_UNKNOWN).sum())
    if late:
        st.error(f"🔴 지연 후보(일정 확인 필요): {late}건 — 완료예정일이 기준일보다 지난 미완료 건")
    if unknown:
        st.warning(f"⚠️ 완료예정일이 비어 있거나 형식이 잘못된 미완료 건: {unknown}건")

    left, right = st.columns(2)
    with left:
        st.subheader("시험분야별 건수")
        st.altair_chart(field_bar_chart(view), width="stretch")
    with right:
        st.subheader("진행상태 비율")
        st.altair_chart(status_donut_chart(view), width="stretch")

    # ───────────────────────── 의뢰 목록 ─────────────────────────
    st.subheader("📄 의뢰 목록")
    table = view[COLUMNS + ["잔여일수", "일정 상태"]].sort_values("접수일")
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

    render_detail(view)


# ───────────────────────── Excel 다운로드 ─────────────────────────
def project_sheet(df):
    """프로젝트목록 시트용 표: 날짜는 문자열로, 빈 잔여일수는 빈 칸으로 변환."""
    out = df[COLUMNS + ["잔여일수", "일정 상태"]].sort_values("접수일").copy()
    for col in ["접수일", "완료예정일"]:
        out[col] = out[col].dt.strftime("%Y-%m-%d")
    out["잔여일수"] = out["잔여일수"].astype(object).where(out["잔여일수"].notna(), None)
    return out


def status_sheet(df):
    """상태별집계 시트용 표 2개: 진행상태별(합계 포함), 일정 상태별 건수·비율."""
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
    """담당자별집계 시트용 표: 담당자 × 진행상태 건수, 합계, 지연 후보 건수."""
    pivot = pd.crosstab(df["담당엔지니어"], df["진행상태"]).reindex(columns=STATUSES, fill_value=0)
    pivot["합계"] = pivot.sum(axis=1)
    pivot["지연 후보"] = (
        df[df["일정 상태"] == SCHEDULE_LATE].groupby("담당엔지니어").size()
        .reindex(pivot.index, fill_value=0)
    )
    return pivot.reset_index().rename(columns={"index": "담당엔지니어"})


def condition_sheet(scope, df, filtered):
    """분석조건 시트용 표: 기준일, 다운로드 일시, 적용한 필터 등을 항목/값으로 기록."""
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
        ("비고", "교육용 가상 데이터. 세션 중 변경한 시험분야(이름 변경·재분류)와 메모가 반영된 결과입니다."),
    ]
    return [pd.DataFrame(rows, columns=["항목", "값"])]


def text_width(value):
    """엑셀 열 너비 계산용 표시 폭 (한글 등 비ASCII 문자는 2칸)."""
    s = "" if value is None else str(value)
    return sum(2 if ord(ch) > 127 else 1 for ch in s)


def write_sheet(writer, name, tables):
    """시트 1행에 교육용 안내 문구를 넣고, 3행부터 표들을 이어 쓰고 머리글·열 너비를 꾸밈."""
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
    """4개 시트(프로젝트목록/상태별집계/담당자별집계/분석조건)로 된 xlsx 바이트를 생성."""
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
    mime=XLSX_MIME,
    width="stretch",
)
d2.download_button(
    f"현재 필터 결과 ({len(view)}건)",
    data=build_excel(view, "현재 필터 결과", filtered=True) if len(view) else b"",
    file_name=f"standardbank_project_status_filtered_{stamp}.xlsx",
    mime=XLSX_MIME,
    disabled=view.empty,
    help="필터 결과가 0건이면 다운로드할 수 없습니다." if view.empty else None,
    type="primary",
    width="stretch",
)
