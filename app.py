import io
import os
import tempfile
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import Workbook

import alliance_stats_report as report


st.set_page_config(page_title="Alliance Stats", layout="wide")


def check_password():
    password = os.environ.get("ALLIANCE_STATS_PASSWORD", "").strip()
    if not password:
        return True

    entered = st.text_input("Password", type="password")
    if entered == password:
        return True
    st.stop()


def save_uploads(uploaded_files, target_dir):
    for uploaded in uploaded_files:
        path = target_dir / uploaded.name
        path.write_bytes(uploaded.getbuffer())


def parse_member_text(text):
    members = set()
    for line in text.splitlines():
        value = line.strip()
        if not value or value.startswith("#"):
            continue
        members.add(value)
    return members or None


def comparison_frame(comparison, source_column):
    rows = []
    for row in comparison:
        rows.append(
            {
                "상태": row["상태"],
                "회원": row[report.KEY_COLUMN],
                "그룹": row[report.GROUP_COLUMN],
                "이전값": row[f"이전 {source_column}"],
                "현재값": row[f"현재 {source_column}"],
                "변화": row[f"{source_column} 변화"],
            }
        )
    return pd.DataFrame(rows).sort_values("변화", ascending=False)


def trend_frame(snapshots, source_column, member_filter):
    rows = []
    for snapshot in snapshots:
        label = snapshot["label"]
        for member, source in snapshot["rows"].items():
            if member_filter and member not in member_filter:
                continue
            rows.append(
                {
                    "날짜": label,
                    "회원": member,
                    "그룹": source.get(report.GROUP_COLUMN, ""),
                    "값": source.get(source_column, 0),
                }
            )

    if member_filter:
        existing = {(row["날짜"], row["회원"]) for row in rows}
        for snapshot in snapshots:
            for member in member_filter:
                key = (snapshot["label"], member)
                if key not in existing:
                    rows.append({"날짜": snapshot["label"], "회원": member, "그룹": "", "값": 0})

    return pd.DataFrame(rows)


def make_excel(snapshots, comparison, top_n, member_filter):
    wb = Workbook()
    report.add_workbook_sheets(wb, snapshots, comparison, top_n, member_filter=member_filter)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def render_metric_tab(title, source_column, snapshots, comparison, member_filter):
    df = comparison_frame(comparison, source_column)
    st.dataframe(df, use_container_width=True, hide_index=True)

    trend = trend_frame(snapshots, source_column, member_filter)
    if not trend.empty:
        fig = px.line(trend, x="날짜", y="값", color="회원", markers=True, title=f"{title} 추이")
        fig.update_layout(height=620, legend_title_text="회원")
        st.plotly_chart(fig, use_container_width=True)


def main():
    check_password()

    st.title("Alliance Stats")

    uploaded_files = st.file_uploader("CSV", type=["csv"], accept_multiple_files=True)
    member_file = st.file_uploader("members.txt", type=["txt"])
    member_text = st.text_area("회원 목록", height=180, placeholder="한 줄에 한 명씩 입력")
    use_all_members = st.toggle("전체 인원", value=False)
    top_n = st.slider("그래프 인원", 5, 100, 30)

    if not uploaded_files:
        st.info("CSV 파일을 업로드하세요.")
        return

    members = None
    if not use_all_members:
        if member_file:
            members = parse_member_text(member_file.getvalue().decode("utf-8-sig"))
        elif member_text.strip():
            members = parse_member_text(member_text)

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        save_uploads(uploaded_files, temp_path)
        snapshots = report.load_snapshots(temp_path)
        comparison = report.build_comparison(snapshots[-2], snapshots[-1], members)

        st.caption(f"{snapshots[-2]['path'].name} -> {snapshots[-1]['path'].name}")

        c1, c2, c3 = st.columns(3)
        c1.metric("CSV", len(snapshots))
        c2.metric("회원", len(comparison))
        c3.metric("필터", "전체" if not members else f"{len(members)}명")

        tabs = st.tabs(["세력치", "공헌", "전공", "협공", "기부", "랭킹"])
        metrics = [
            ("세력치", "세력치"),
            ("공헌", "공헌총량"),
            ("전공", "전공총량"),
            ("협공", "협공총량"),
            ("기부", "기부총량"),
            ("랭킹", "공헌 랭킹"),
        ]

        for tab, (title, source_column) in zip(tabs, metrics):
            with tab:
                render_metric_tab(title, source_column, snapshots, comparison, members)

        excel = make_excel(snapshots, comparison, top_n, members)
        st.download_button(
            "엑셀 다운로드",
            data=excel,
            file_name="alliance_stats_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
