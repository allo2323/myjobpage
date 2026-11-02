import hashlib
import html
from datetime import date, datetime
from io import BytesIO, StringIO
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="스탠다드 뱅크 인증 프로젝트 대시보드",
    page_icon="📋",
    layout="wide",
)

DATA_DIR = Path(__file__).resolve().parent / "data"
# 과제 안내의 파일명과 실제 배포 파일명(s 없음)을 모두 찾는다.
PROJECT_FILE_NAMES = ("SB_projects_demo.csv", "SB_project_demo.csv")
RESOURCE_FILE_NAMES = ("SB_resources_demo.csv", "SB_resource_demo.csv")

DATA_COLUMNS = [
    "의뢰번호",
    "고객사",
    "제품명",
    "모델명",
    "시험분야",
    "적용규격",
    "담당엔지니어",
    "접수일",
    "완료예정일",
    "진행상태",
    "시료상태",
    "메모",
]
RESOURCE_COLUMNS = ["자원", "상태", "비고"]
PROJECT_STATUSES = [
    "접수",
    "자료확인",
    "시험진행",
    "성적서검토",
    "보완요청",
    "완료",
]
DEFAULT_TEST_FIELDS = ["EMC", "무선(RF)", "인증서류", "SAR"]
FILTER_COLUMNS = ["시험분야", "적용규격", "담당엔지니어", "진행상태", "시료상태"]

SCHEDULE_DONE = "완료"
SCHEDULE_OVERDUE = "지연 후보(일정 확인 필요)"
SCHEDULE_DUE_TODAY = "오늘 마감"
SCHEDULE_UPCOMING = "진행 예정"
SCHEDULE_UNKNOWN = "일정 확인 필요"
SCHEDULE_STATUSES = [
    SCHEDULE_OVERDUE,
    SCHEDULE_DUE_TODAY,
    SCHEDULE_UPCOMING,
    SCHEDULE_UNKNOWN,
    SCHEDULE_DONE,
]
SCHEDULE_RULE_TEXT = (
    "완료 → '완료'(예정일 경과와 무관) / 완료예정일 없음·형식 오류 → '일정 확인 필요' / "
    "잔여일수 < 0 → '지연 후보' / 잔여일수 = 0 → '오늘 마감' / 그 외 → '진행 예정'"
)

# 진행 단계 순서대로 옅은 파랑 → 진한 네이비, 보완요청만 주의색
STATUS_COLORS = ["#bfdbfe", "#93c5fd", "#3b82f6", "#1d4ed8", "#f59e0b", "#102a56"]
# 자원 상태별 (글자색, 배경색). 정의되지 않은 상태는 회색으로 그대로 표시한다.
RESOURCE_STATUS_STYLES = {
    "사용 가능": ("#15803d", "#dcfce7"),
    "사용 중": ("#1d4ed8", "#dbeafe"),
    "점검 중": ("#b45309", "#fef3c7"),
}
RESOURCE_UNKNOWN_STYLE = ("#475569", "#f1f5f9")

NOTICE_TEXT = "교육용 가상 데이터 기준 — 실제 업무 시스템 아님"
SESSION_NOTICE_TEXT = (
    "시험분야 관리·메모 수정 등 화면에서 바꾼 내용은 현재 브라우저 세션(session_state)에만 "
    "보관됩니다. 새로고침하거나 세션이 종료되면 사라질 수 있으며, 원본 CSV 파일은 변경되지 않습니다."
)
RESOURCE_NOTICE_TEXT = "현황 조회용이며 예약·시험 재개를 승인하는 기능이 아님"
EXCEL_NOTICE_TEXT = "교육용 검토 파일 — 공식 보고서 아님"
EMPTY_RESULT_TEXT = "조건에 맞는 인증 프로젝트가 없습니다"

st.markdown(
    f"""
    <style>
    .stApp {{
        background: #ffffff;
        color: #172b4d;
    }}
    [data-testid="stHeader"] {{
        background: rgba(255, 255, 255, 0.95);
    }}
    [data-testid="stSidebar"] {{
        background: #f4f7fb;
    }}
    .notice-banner {{
        padding: 0.7rem 1rem;
        margin: 0.1rem 0 0.6rem;
        border: 1px solid #bfdbfe;
        border-radius: 0.55rem;
        background: #eff6ff;
        color: #1e3a8a;
        font-weight: 600;
    }}
    .metric-card {{
        height: 100%;
        min-height: 96px;
        padding: 0.9rem 1rem;
        border: 1px solid #dbe5f0;
        border-top: 4px solid #2563eb;
        border-radius: 0.65rem;
        background: #ffffff;
        box-shadow: 0 2px 8px rgba(23, 43, 77, 0.05);
    }}
    .metric-card.alert {{
        border-top-color: #dc2626;
        background: #fef2f2;
    }}
    .metric-card.alert .metric-value {{
        color: #b91c1c;
    }}
    .metric-label {{
        color: #52647a;
        font-size: 0.85rem;
        font-weight: 600;
    }}
    .metric-value {{
        margin-top: 0.35rem;
        color: #102a56;
        font-size: 1.7rem;
        font-weight: 700;
        line-height: 1.2;
    }}
    .resource-card {{
        padding: 1rem;
        border: 1px solid #dbe5f0;
        border-left: 5px solid var(--status-color);
        border-radius: 0.65rem;
        background: #ffffff;
    }}
    .resource-name {{
        color: #102a56;
        font-size: 1.05rem;
        font-weight: 700;
    }}
    .resource-badge {{
        display: inline-block;
        margin: 0.5rem 0 0.4rem;
        padding: 0.15rem 0.6rem;
        border-radius: 999px;
        background: var(--status-bg);
        color: var(--status-color);
        font-weight: 700;
    }}
    .resource-note, .resource-disclaimer {{
        color: #52647a;
        font-size: 0.82rem;
    }}
    .resource-disclaimer {{
        margin-top: 0.5rem;
        padding-top: 0.4rem;
        border-top: 1px dashed #dbe5f0;
    }}
    .detail-grid {{
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 0.6rem;
    }}
    .detail-item {{
        padding: 0.6rem 0.8rem;
        border: 1px solid #dbe5f0;
        border-radius: 0.5rem;
        background: #f8fafc;
    }}
    .detail-label {{
        color: #52647a;
        font-size: 0.8rem;
        font-weight: 600;
    }}
    .detail-value {{
        color: #102a56;
        font-weight: 600;
    }}
    h1, h2, h3 {{
        color: #102a56;
    }}
    </style>
    <div class="notice-banner">{NOTICE_TEXT}</div>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# 데이터 읽기
# ---------------------------------------------------------------------------

# 업로드 파일이 있으면 그것을, 없으면 data 폴더에서 file_names 순서대로 찾은 파일을
# (원본 바이트, 표시용 파일명)으로 돌려준다. 어느 쪽도 없으면 None.
def find_csv_source(uploaded_file, file_names):
    if uploaded_file is not None:
        return uploaded_file.getvalue(), uploaded_file.name
    for name in file_names:
        path = DATA_DIR / name
        if path.is_file():
            try:
                return path.read_bytes(), name
            except OSError as error:
                st.error(f"{name} 파일을 읽지 못했습니다: {error}")
                return None
    return None


# CSV 바이트를 UTF-8 또는 CP949로 해석해 문자열 DataFrame으로 읽고 필수 열을 검사한다.
# 문제가 있으면 화면에 오류를 표시하고 None을 돌려준다.
def parse_csv(csv_bytes, required_columns, label):
    try:
        try:
            csv_text = csv_bytes.decode("utf-8-sig")
        except UnicodeDecodeError:
            csv_text = csv_bytes.decode("cp949")
        frame = pd.read_csv(StringIO(csv_text), dtype=str)
    except UnicodeDecodeError:
        st.error(f"{label} CSV 인코딩을 읽을 수 없습니다. UTF-8 또는 CP949 형식으로 저장해 주세요.")
        return None
    except (pd.errors.EmptyDataError, pd.errors.ParserError) as error:
        st.error(f"{label} CSV 형식을 읽을 수 없습니다: {error}")
        return None

    frame.columns = frame.columns.astype(str).str.strip()
    missing_columns = [column for column in required_columns if column not in frame.columns]
    if missing_columns:
        st.error(f"{label} CSV에 필수 항목이 없습니다: {', '.join(missing_columns)}")
        return None
    frame = frame[required_columns].copy()
    for column in required_columns:
        frame[column] = frame[column].str.strip()
    return frame


# 데이터 출처(파일 내용)가 바뀌었을 때만 session_state의 작업용 데이터를 새로 만든다.
# 같은 파일이면 시험분야 변경·메모 수정 내용을 그대로 유지한다.
def init_project_state(projects, csv_bytes):
    source_key = hashlib.md5(csv_bytes).hexdigest()
    if st.session_state.get("project_source_key") == source_key:
        return
    st.session_state.project_source_key = source_key
    st.session_state.projects = projects.copy()
    test_fields = list(DEFAULT_TEST_FIELDS)
    for field in projects["시험분야"].dropna().unique():
        if field not in test_fields:
            test_fields.append(field)
    st.session_state.test_fields = test_fields
    for column in FILTER_COLUMNS:
        st.session_state.pop(f"filter_{column}", None)


# 정의된 진행상태 순서에, 데이터에만 있는 미정의 상태를 뒤에 붙인 목록을 돌려준다.
def status_options(projects):
    extras = sorted(set(projects["진행상태"].dropna()) - set(PROJECT_STATUSES))
    return PROJECT_STATUSES + extras


# 기준일 대비 잔여일수·일정상태 열을 완료예정일 바로 뒤에 추가한 사본을 돌려준다.
# 판정 우선순위: 완료 > 날짜 없음/오류 > 지연 후보 > 오늘 마감 > 진행 예정
def add_schedule_columns(projects, as_of_date):
    due_dates = pd.to_datetime(projects["완료예정일"], errors="coerce")
    remaining_days = (due_dates - pd.Timestamp(as_of_date)).dt.days.astype("Int64")

    schedule = pd.Series(SCHEDULE_UPCOMING, index=projects.index, dtype=object)
    schedule = schedule.mask(remaining_days.eq(0).fillna(False).astype(bool), SCHEDULE_DUE_TODAY)
    schedule = schedule.mask(remaining_days.lt(0).fillna(False).astype(bool), SCHEDULE_OVERDUE)
    schedule = schedule.mask(due_dates.isna(), SCHEDULE_UNKNOWN)
    schedule = schedule.mask(projects["진행상태"].eq("완료"), SCHEDULE_DONE)

    result = projects.copy()
    position = result.columns.get_loc("완료예정일") + 1
    result.insert(position, "잔여일수", remaining_days)
    result.insert(position + 1, "일정상태", schedule)
    return result


# ---------------------------------------------------------------------------
# 사이드바: 필터
# ---------------------------------------------------------------------------

# 사이드바에 다중 선택 필터를 그리고 {열 이름: 선택값 목록}을 돌려준다.
# 아무것도 고르지 않은 항목은 '전체'로 본다.
def render_filters(projects):
    selections = {}
    for column in FILTER_COLUMNS:
        if column == "시험분야":
            options = st.session_state.test_fields
        elif column == "진행상태":
            options = status_options(projects)
        else:
            options = sorted(projects[column].dropna().unique())
        selections[column] = st.sidebar.multiselect(
            column, options, key=f"filter_{column}", placeholder="전체"
        )
    return selections


# 선택된 필터 조건을 모두 만족하는 행만 돌려준다.
def apply_filters(projects, selections):
    mask = pd.Series(True, index=projects.index)
    for column, values in selections.items():
        if values:
            mask &= projects[column].isin(values)
    return projects[mask]


# ---------------------------------------------------------------------------
# 사이드바: 카테고리(시험분야) 관리
# 변경은 on_click 콜백에서 처리한다. 콜백은 다음 실행 전에 돌기 때문에
# 이미 그려진 필터 위젯의 상태(filter_시험분야)도 안전하게 함께 고칠 수 있다.
# ---------------------------------------------------------------------------

# 처리 결과 메시지를 다음 화면에서 한 번 보여주도록 저장한다.
def set_flash(area, kind, message):
    st.session_state[f"flash_{area}"] = (kind, message)


# 저장된 처리 결과 메시지를 꺼내 표시한다 (한 번 표시 후 삭제).
def show_flash(area, container=st):
    flash = st.session_state.pop(f"flash_{area}", None)
    if flash:
        kind, message = flash
        getattr(container, kind)(message)


# 해당 시험분야를 쓰는 의뢰 건수를 돌려준다.
def field_usage(field):
    return int(st.session_state.projects["시험분야"].eq(field).sum())


# 시험분야 필터에 선택돼 있던 값을 새 이름으로 바꾸거나(new=None이면) 제거한다.
def update_field_filter(old, new=None):
    key = "filter_시험분야"
    selected = st.session_state.get(key, [])
    if old not in selected:
        return
    updated = [new if value == old else value for value in selected if new or value != old]
    st.session_state[key] = list(dict.fromkeys(updated))


def add_field_callback():
    name = st.session_state.new_field_name.strip()
    if not name:
        set_flash("category", "error", "등록할 시험분야 이름을 입력해 주세요.")
    elif name in st.session_state.test_fields:
        set_flash("category", "warning", f"'{name}'은(는) 이미 등록된 시험분야입니다.")
    else:
        st.session_state.test_fields.append(name)
        st.session_state.new_field_name = ""
        set_flash("category", "success", f"시험분야 '{name}'을(를) 등록했습니다.")


def rename_field_callback():
    old = st.session_state.rename_field_target
    new = st.session_state.rename_field_new.strip()
    fields = st.session_state.test_fields
    if old is None or not new:
        set_flash("category", "error", "변경할 시험분야와 새 이름을 입력해 주세요.")
        return
    if new == old:
        set_flash("category", "info", "기존 이름과 같습니다.")
        return
    if new in fields:
        set_flash("category", "warning", f"'{new}'은(는) 이미 있는 시험분야입니다. 다른 이름을 입력해 주세요.")
        return
    projects = st.session_state.projects
    changed = int(projects["시험분야"].eq(old).sum())
    projects.loc[projects["시험분야"].eq(old), "시험분야"] = new
    fields[fields.index(old)] = new
    update_field_filter(old, new)
    st.session_state.rename_field_target = new
    st.session_state.rename_field_new = ""
    set_flash("category", "success", f"'{old}' → '{new}'로 변경했고, 의뢰 {changed}건도 함께 갱신했습니다.")


def reassign_field_callback():
    source = st.session_state.delete_field_target
    target = st.session_state.reassign_field_target
    projects = st.session_state.projects
    moved = int(projects["시험분야"].eq(source).sum())
    projects.loc[projects["시험분야"].eq(source), "시험분야"] = target
    set_flash("category", "success", f"'{source}' 의뢰 {moved}건을 '{target}'(으)로 재분류했습니다. 이제 삭제할 수 있습니다.")


def delete_field_callback():
    field = st.session_state.delete_field_target
    usage = field_usage(field)
    if usage:
        set_flash("category", "error", f"'{field}'은(는) {usage}건에서 사용 중이라 삭제할 수 없습니다.")
        return
    st.session_state.test_fields.remove(field)
    update_field_filter(field)
    st.session_state.pop("delete_field_target", None)
    st.session_state.pop("rename_field_target", None)
    set_flash("category", "success", f"시험분야 '{field}'을(를) 삭제했습니다.")


# 사이드바에 시험분야 목록과 등록/이름 변경/삭제 메뉴를 그린다.
def render_category_manager():
    sidebar = st.sidebar
    sidebar.subheader("시험분야")
    show_flash("category", sidebar)
    fields = st.session_state.test_fields
    sidebar.dataframe(
        pd.DataFrame({"시험분야": fields, "사용 건수": [field_usage(field) for field in fields]}),
        hide_index=True,
        width="stretch",
    )

    action = sidebar.radio("작업", ["등록", "이름 변경", "삭제"], horizontal=True, key="category_action")
    if action == "등록":
        with sidebar.form("add_field_form", border=False):
            st.text_input("새 시험분야", key="new_field_name")
            st.form_submit_button("등록", on_click=add_field_callback)
    elif action == "이름 변경":
        if not fields:
            sidebar.info("등록된 시험분야가 없습니다.")
            return
        with sidebar.form("rename_field_form", border=False):
            st.selectbox("변경할 시험분야", fields, key="rename_field_target")
            st.text_input("새 이름", key="rename_field_new")
            st.form_submit_button("이름 변경", on_click=rename_field_callback)
        sidebar.caption("이름을 바꾸면 해당 시험분야의 의뢰 데이터도 함께 바뀝니다.")
    else:
        if not fields:
            sidebar.info("등록된 시험분야가 없습니다.")
            return
        target = sidebar.selectbox("삭제할 시험분야", fields, key="delete_field_target")
        usage = field_usage(target)
        if usage:
            sidebar.warning(
                f"'{target}'은(는) 의뢰 {usage}건에서 사용 중이라 바로 삭제할 수 없습니다. "
                "아래에서 다른 시험분야로 재분류한 뒤 삭제해 주세요."
            )
            others = [field for field in fields if field != target]
            if others:
                sidebar.selectbox("재분류할 시험분야", others, key="reassign_field_target")
                sidebar.button(f"{usage}건 재분류", on_click=reassign_field_callback)
            else:
                sidebar.info("재분류할 다른 시험분야가 없습니다. 먼저 새 시험분야를 등록해 주세요.")
        sidebar.button("삭제", type="primary", disabled=bool(usage), on_click=delete_field_callback)


# ---------------------------------------------------------------------------
# 본문 화면 구성
# ---------------------------------------------------------------------------

# (라벨, 값) 목록을 요약 카드 한 줄로 그린다. alert_labels의 카드는 값이 1 이상이면 빨간색으로 강조.
def render_metric_cards(items, alert_labels=()):
    columns = st.columns(len(items))
    for column, (label, value) in zip(columns, items):
        card_class = "metric-card alert" if label in alert_labels and value else "metric-card"
        column.markdown(
            f"""
            <div class="{card_class}">
                <div class="metric-label">{html.escape(label)}</div>
                <div class="metric-value">{value:,}건</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# 시험분야별 건수 막대그래프와 진행상태 비율 도넛 차트를 나란히 그린다. (0건이면 호출하지 않는다)
def render_charts(projects):
    field_counts = (
        projects["시험분야"].fillna("미지정").value_counts().rename_axis("시험분야").reset_index(name="건수")
    )
    bar_chart = (
        alt.Chart(field_counts)
        .mark_bar(color="#1e3a8a", cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
        .encode(
            x=alt.X("시험분야:N", sort="-y", axis=alt.Axis(labelAngle=0, title=None)),
            y=alt.Y("건수:Q", axis=alt.Axis(tickMinStep=1, title="건수")),
            tooltip=["시험분야", "건수"],
        )
        .properties(height=300)
    )

    statuses = status_options(projects)
    status_counts = projects["진행상태"].value_counts().reindex(statuses, fill_value=0)
    status_counts = status_counts[status_counts > 0].rename_axis("진행상태").reset_index(name="건수")
    status_counts["비율"] = status_counts["건수"] / status_counts["건수"].sum()
    colors = (STATUS_COLORS + ["#94a3b8"] * len(statuses))[: len(statuses)]
    donut_chart = (
        alt.Chart(status_counts)
        .mark_arc(innerRadius=70, stroke="#ffffff", strokeWidth=2)
        .encode(
            theta=alt.Theta("건수:Q"),
            color=alt.Color(
                "진행상태:N",
                scale=alt.Scale(domain=statuses, range=colors),
                sort=statuses,
                legend=alt.Legend(title=None, orient="right"),
            ),
            order=alt.Order("order:Q"),
            tooltip=["진행상태", "건수", alt.Tooltip("비율:Q", format=".0%")],
        )
        .transform_calculate(order=f"indexof({statuses!r}, datum['진행상태'])")
        .properties(height=300)
    )

    bar_column, donut_column = st.columns(2)
    with bar_column:
        st.markdown("**시험분야별 건수**")
        st.altair_chart(bar_chart, width="stretch")
    with donut_column:
        st.markdown("**진행상태 비율**")
        st.altair_chart(donut_chart, width="stretch")


# 지연 후보 행을 빨간색 계열로 칠하고 빈 값은 공란으로 보이게 한 Styler를 돌려준다.
def style_projects(projects):
    def row_style(row):
        if row["일정상태"] == SCHEDULE_OVERDUE:
            return ["background-color: #fee2e2; color: #991b1b; font-weight: 600"] * len(row)
        return [""] * len(row)

    return projects.style.apply(row_style, axis=1).format(na_rep="")


# 의뢰번호를 고르면 상세 정보와 메모 수정 폼을 보여준다.
def render_project_detail(projects):
    labels = {
        row["의뢰번호"]: f"{row['의뢰번호']} · {row['고객사']} · {row['제품명']}"
        for _, row in projects.iterrows()
    }
    project_id = st.selectbox("의뢰번호 선택", list(labels), format_func=labels.get)
    row = projects[projects["의뢰번호"].eq(project_id)].iloc[0]

    def text(value):
        return "-" if pd.isna(value) or value == "" else str(value)

    due_text = text(row["완료예정일"])
    if pd.notna(row["잔여일수"]):
        due_text += f" (잔여 {row['잔여일수']}일)"
    details = [
        ("고객사", text(row["고객사"])),
        ("제품 · 모델", f"{text(row['제품명'])} / {text(row['모델명'])}"),
        ("시험분야", text(row["시험분야"])),
        ("적용규격", text(row["적용규격"])),
        ("담당엔지니어", text(row["담당엔지니어"])),
        ("시료상태", text(row["시료상태"])),
        ("완료예정일", due_text),
        ("진행상태 / 일정상태", f"{text(row['진행상태'])} / {row['일정상태']}"),
    ]
    items_html = "".join(
        f'<div class="detail-item"><div class="detail-label">{html.escape(label)}</div>'
        f'<div class="detail-value">{html.escape(value)}</div></div>'
        for label, value in details
    )
    st.markdown(f'<div class="detail-grid">{items_html}</div>', unsafe_allow_html=True)

    show_flash("memo")
    memo_key = f"memo_input_{project_id}"
    with st.form("memo_form", border=False):
        st.text_area("담당자 메모", value="" if pd.isna(row["메모"]) else row["메모"], key=memo_key)
        st.form_submit_button("메모 저장", on_click=save_memo_callback, args=(project_id, memo_key))


# 입력창의 메모를 session_state의 의뢰 데이터에 저장한다.
def save_memo_callback(project_id, memo_key):
    projects = st.session_state.projects
    projects.loc[projects["의뢰번호"].eq(project_id), "메모"] = st.session_state[memo_key].strip()
    set_flash("memo", "success", f"{project_id} 메모를 저장했습니다. (현재 세션에만 보관)")


# 자원 CSV의 상태를 카드로 보여준다. 상태 값은 파일 내용 그대로 표시하며,
# '점검 중' 등 어떤 상태도 코드에서 바꾸지 않는다.
def render_resource_cards(resources):
    if resources.empty:
        st.info("자원 데이터가 비어 있습니다.")
        return
    columns = st.columns(len(resources))
    for column, (_, row) in zip(columns, resources.iterrows()):
        status = row["상태"] if pd.notna(row["상태"]) else "상태 미입력"
        text_color, background = RESOURCE_STATUS_STYLES.get(status, RESOURCE_UNKNOWN_STYLE)
        note = "" if pd.isna(row["비고"]) else row["비고"]
        column.markdown(
            f"""
            <div class="resource-card" style="--status-color: {text_color}; --status-bg: {background};">
                <div class="resource-name">{html.escape(str(row["자원"]))}</div>
                <div class="resource-badge">{html.escape(status)}</div>
                <div class="resource-note">{html.escape(note)}</div>
                <div class="resource-disclaimer">{RESOURCE_NOTICE_TEXT}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ---------------------------------------------------------------------------
# Excel 다운로드
# ---------------------------------------------------------------------------

# 진행상태·일정상태별 건수와 비율 표를 만든다.
def summarize_by_status(projects):
    total = len(projects)
    rows = []
    for group, column, statuses in (
        ("진행상태", "진행상태", status_options(projects)),
        ("일정상태", "일정상태", SCHEDULE_STATUSES),
    ):
        for status in statuses:
            count = int(projects[column].eq(status).sum())
            rows.append((group, status, count, round(count / total * 100, 1) if total else 0.0))
    return pd.DataFrame(rows, columns=["구분", "상태", "건수", "비율(%)"])


# 담당엔지니어별 전체·진행상태별·지연 후보 건수 표를 만든다.
def summarize_by_engineer(projects):
    statuses = status_options(projects)
    engineers = projects["담당엔지니어"].fillna("미지정")
    table = pd.crosstab(engineers, projects["진행상태"]).reindex(columns=statuses, fill_value=0)
    table.insert(0, "전체", table.sum(axis=1))
    overdue = projects["일정상태"].eq(SCHEDULE_OVERDUE).groupby(engineers).sum()
    table["지연 후보"] = overdue.reindex(table.index, fill_value=0).astype(int)
    table.columns.name = None
    return table.rename_axis("담당엔지니어").reset_index()


# 분석조건 시트 내용(범위, 기준일, 필터)을 만든다. selections가 None이면 전체 목록.
def build_conditions(scope, as_of_date, selections, row_count, data_source):
    conditions = {
        "다운로드 범위": scope,
        "기준일": f"{as_of_date:%Y-%m-%d}",
        "생성 일시": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
        "데이터 파일": data_source,
        "대상 건수": f"{row_count}건",
    }
    for column in FILTER_COLUMNS:
        if selections is None:
            value = "적용 안 함(전체 목록)"
        else:
            value = ", ".join(selections[column]) if selections[column] else "전체(선택 없음)"
        conditions[f"필터 - {column}"] = value
    conditions["일정상태 판단 기준"] = SCHEDULE_RULE_TEXT
    conditions["비고"] = SESSION_NOTICE_TEXT
    return pd.DataFrame(list(conditions.items()), columns=["항목", "값"])


# 프로젝트목록/상태별집계/담당자별집계/분석조건 4개 시트의 xlsx 바이트를 만든다.
# 각 시트 1행에 교육용 안내 문구를 넣고, 표는 3행부터 시작한다.
def build_excel(projects, conditions):
    sheets = {
        "프로젝트목록": projects,
        "상태별집계": summarize_by_status(projects),
        "담당자별집계": summarize_by_engineer(projects),
        "분석조건": conditions,
    }
    header_fill = PatternFill("solid", fgColor="102A56")
    overdue_fill = PatternFill("solid", fgColor="FEE2E2")
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)
            sheet = writer.sheets[sheet_name]
            sheet["A1"] = EXCEL_NOTICE_TEXT
            sheet["A1"].font = Font(bold=True, color="B91C1C", size=12)
            for cell in sheet[3]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = header_fill
            for index, column in enumerate(frame.columns, start=1):
                lengths = [len(str(column))] + [len(str(value)) for value in frame[column].tolist()]
                sheet.column_dimensions[get_column_letter(index)].width = min(max(lengths) * 1.8 + 2, 80)
            if "일정상태" in frame.columns:
                for offset, status in enumerate(frame["일정상태"].tolist()):
                    if status == SCHEDULE_OVERDUE:
                        for cell in sheet[4 + offset]:
                            cell.fill = overdue_fill
            sheet.freeze_panes = "A4"
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# 화면 실행
# ---------------------------------------------------------------------------

st.info(SESSION_NOTICE_TEXT, icon="ℹ️")

title_column, date_column = st.columns([4, 1])
with title_column:
    st.title("스탠다드 뱅크 인증 프로젝트")
    st.caption("프로젝트 접수 및 진행현황 대시보드")
with date_column:
    as_of_date = st.date_input("기준일", value=date(2026, 10, 23))

uploaded_file = st.file_uploader(
    "프로젝트 CSV 업로드",
    type=["csv"],
    help="data/SB_projects_demo.csv가 없거나 다른 파일을 사용하려면 CSV를 업로드하세요.",
)
project_source = find_csv_source(uploaded_file, PROJECT_FILE_NAMES)
if project_source is None:
    st.info(
        "프로젝트 데이터 파일을 찾을 수 없습니다. "
        "`data/SB_projects_demo.csv`를 추가하거나 위에서 CSV 파일을 업로드해 주세요."
    )
    st.stop()
project_bytes, data_source = project_source
loaded_projects = parse_csv(project_bytes, DATA_COLUMNS, "프로젝트")
if loaded_projects is None:
    st.stop()
init_project_state(loaded_projects, project_bytes)
st.caption(f"데이터 파일: {data_source}")

projects = st.session_state.projects
unknown_statuses = sorted(set(projects["진행상태"].dropna()) - set(PROJECT_STATUSES))
if unknown_statuses:
    st.warning(f"정의되지 않은 진행상태가 있습니다: {', '.join(unknown_statuses)}")

st.sidebar.header("필터")
st.sidebar.caption("선택하지 않은 항목은 전체로 봅니다.")
selections = render_filters(projects)
st.sidebar.divider()
st.sidebar.header("카테고리 관리")
render_category_manager()

all_projects = add_schedule_columns(projects, as_of_date)
filtered = apply_filters(all_projects, selections)
active_filters = [column for column in FILTER_COLUMNS if selections[column]]

invalid_due_dates = int((all_projects["완료예정일"].notna() & all_projects["잔여일수"].isna()).sum())
if invalid_due_dates:
    st.warning(f"완료예정일을 날짜로 읽지 못한 항목이 {invalid_due_dates}건 있습니다. ('일정 확인 필요'로 표시)")

st.subheader("요약")
st.caption(
    f"기준일 {as_of_date:%Y-%m-%d} · "
    + (f"필터 적용({', '.join(active_filters)}) {len(filtered)}건 / 전체 {len(all_projects)}건"
       if active_filters else f"전체 {len(all_projects)}건")
)
render_metric_cards(
    [("조회 의뢰", len(filtered))]
    + [(status, int(filtered["진행상태"].eq(status).sum())) for status in PROJECT_STATUSES]
)
st.write("")
render_metric_cards(
    [(status, int(filtered["일정상태"].eq(status).sum())) for status in SCHEDULE_STATUSES],
    alert_labels={SCHEDULE_OVERDUE},
)
st.caption(f"일정상태 기준: {SCHEDULE_RULE_TEXT}")

if filtered.empty:
    st.info(EMPTY_RESULT_TEXT)
else:
    st.subheader("현황 차트")
    render_charts(filtered)

    overdue_projects = filtered[filtered["일정상태"].eq(SCHEDULE_OVERDUE)]
    st.subheader(f"지연 후보 의뢰 ({len(overdue_projects)}건)")
    if overdue_projects.empty:
        st.success("기준일 기준 지연 후보 의뢰가 없습니다.")
    else:
        overdue_view = overdue_projects.assign(지연일수=-overdue_projects["잔여일수"])
        st.dataframe(
            overdue_view.sort_values("지연일수", ascending=False)[
                ["지연일수", "의뢰번호", "고객사", "제품명", "담당엔지니어",
                 "완료예정일", "진행상태", "메모"]
            ],
            width="stretch",
            hide_index=True,
        )

    st.subheader(f"의뢰 목록 ({len(filtered)}건)")
    st.dataframe(style_projects(filtered), width="stretch", hide_index=True)

    st.subheader("의뢰 상세 조회")
    render_project_detail(filtered)

st.subheader("Excel 다운로드")
download_date = f"{date.today():%Y%m%d}"
full_column, filtered_column = st.columns(2)
with full_column:
    st.download_button(
        f"전체 목록 다운로드 ({len(all_projects)}건)",
        data=build_excel(
            all_projects,
            build_conditions("전체 목록", as_of_date, None, len(all_projects), data_source),
        ),
        file_name=f"standardbank_project_status_{download_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )
with filtered_column:
    st.download_button(
        f"현재 필터 결과 다운로드 ({len(filtered)}건)",
        data=build_excel(
            filtered,
            build_conditions("현재 필터 결과", as_of_date, selections, len(filtered), data_source),
        ) if not filtered.empty else b"",
        file_name=f"standardbank_project_status_filtered_{download_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=filtered.empty,
        width="stretch",
    )
st.caption(f"시트: 프로젝트목록 / 상태별집계 / 담당자별집계 / 분석조건 · 각 시트 상단에 '{EXCEL_NOTICE_TEXT}' 표시")

st.subheader("시험 자원 현황")
st.caption(RESOURCE_NOTICE_TEXT + " · 상태는 CSV 내용을 그대로 표시합니다.")
with st.expander("자원 CSV 업로드"):
    uploaded_resource_file = st.file_uploader(
        "자원 CSV",
        type=["csv"],
        help="data/SB_resources_demo.csv가 없거나 다른 파일을 사용하려면 업로드하세요.",
        key="resource_uploader",
    )
resource_source = find_csv_source(uploaded_resource_file, RESOURCE_FILE_NAMES)
if resource_source is None:
    st.info(
        "자원 데이터 파일을 찾을 수 없습니다. "
        "`data/SB_resources_demo.csv`를 추가하거나 위의 '자원 CSV 업로드'에서 올려 주세요."
    )
else:
    resources = parse_csv(resource_source[0], RESOURCE_COLUMNS, "자원")
    if resources is not None:
        render_resource_cards(resources)
        st.caption(f"자원 데이터 파일: {resource_source[1]}")
