from __future__ import annotations
import os
import tempfile
import yaml
import streamlit as st
from dotenv import load_dotenv
from sources import excel_reader, srs_reader, terra_reader, jira_client
from filler import fill_sstr

load_dotenv()
st.set_page_config(page_title="SSTR Auto-Fill", page_icon="📄", layout="centered")
st.title("📄 SSTR 자동 채움")

with open("config.yaml", encoding="utf-8") as f:
    cfg = yaml.safe_load(f)


def _save(up):
    if up is None:
        return None
    t = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(up.name)[1])
    t.write(up.getbuffer())
    t.close()
    return t.name


st.header("1. 파일 업로드")
c1, c2 = st.columns(2)
terra_path = _save(c1.file_uploader("teRA 추출본 (템플릿)", type=["docx"]))
excel_path = _save(c2.file_uploader("eepxtool 엑셀", type=["xls", "xlsx"]))
srs_path = _save(st.file_uploader("SRS", type=["docx"]))

variant_info = None
if excel_path:
    st.header("2. 차종 선택")
    try:
        variants = excel_reader.load_variants(excel_path, cfg)
        if variants:
            code = st.selectbox("Variant", sorted(variants))
            variant_info = variants[code]
            a, b, c = st.columns(3)
            a.metric("Customer", variant_info["customer"])
            b.metric("Model", variant_info["model"])
            c.metric("Part Number", variant_info["part_number"])
        else:
            st.warning("Variant_1 에서 변형을 못 찾음.")
    except Exception as e:
        st.error(f"엑셀 읽기 실패: {e}")

sw_version_full = None
if srs_path and variant_info:
    st.header("3. SW 버전 (SRS)")
    phase = st.selectbox(
        "Phase", ["Proto", "M/CAR", "P1", "P2", "M", "AP1", "AP2", "R/C", "SOP"]
    )
    vehicle = variant_info["model"].replace("PE", " PE").strip()
    sw_version_full = srs_reader.find_sw_version(srs_path, vehicle, phase)
    if sw_version_full:
        st.success(f"SW Version: {sw_version_full}")
    else:
        sw_version_full = st.text_input("SW Version (수동)")
    rev = srs_reader.load_latest_revision(srs_path)
    if rev:
        st.caption(f"SRS 최신 리비전: {rev['ver']} ({rev['released']})")

st.header("4. 선택 항목")
c1, c2 = st.columns(2)
test_type = c1.selectbox("Test Type", cfg["choices"]["test_type"])
test_result = c2.selectbox("Test Result", cfg["choices"]["test_result"])
principal = st.text_input("Principal (PM, 영문)")
hw_version = st.text_input("HW-Version", "1.00")
pcb_index = st.text_input("PCB-Index", "N/A")

st.header("5. Jira (선택)")
objective_key = st.text_input("Objective 이슈 키 (예: KHL-9875)")
jira_project = st.text_input("Jira 프로젝트", "KKZ")
fix_version = st.text_input("fixVersion", sw_version_full or "")
if st.button("Jira 조회") and objective_key:
    try:
        jc = jira_client.connect(cfg)
        st.write(jira_client.get_objective(jc, objective_key))
        if jira_project and fix_version:
            issues = jira_client.issues_by_fixversion(jc, jira_project, fix_version)
            st.write(f"fixVersion 매칭 {len(issues)}건")
            st.dataframe(issues)
    except Exception as e:
        st.error(f"Jira 오류: {e}")

st.header("6. SSTR 생성")
if st.button("자동 채움 실행", type="primary"):
    if not (terra_path and variant_info):
        st.error("teRA 템플릿 + 엑셀(차종) 필수.")
    else:
        terra = terra_reader.load_terra_fields(terra_path)
        sw_full = (
            (terra.get("sw_version") or "")
            + ("\n" + sw_version_full if sw_version_full else "")
        ).strip()
        data = {
            "Customer": variant_info["customer"],
            "Project Designation: (Project/Type)": variant_info["model"].replace("PE", " PE"),
            "Principal: (Name / Department.)": principal,
            "HW-Version": hw_version,
            "Part Number\n(Customer/AUMOVIO)":
                f'{variant_info["part_number"]}({variant_info["model"]})',
            "SW-Version": sw_full,
            "PCB-Index": pcb_index,
            "Discipline": cfg["fixed"]["discipline"],
            "Test Level": cfg["fixed"]["test_level"],
            "Test Type": test_type,
            "Test Result": test_result,
            "Result": test_result,
        }
        out = os.path.join(tempfile.gettempdir(), "SSTR_filled.docx")
        fill_sstr(terra_path, out, data, terra=terra)  # B 로직 전달
        with open(out, "rb") as f:
            st.download_button("📥 완성본 다운로드", f, file_name="SSTR_filled.docx")
        st.success("생성 완료 (Metrics·Test Matrix·Period 복사 포함)")
