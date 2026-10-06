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

# -- 3. SW version (concept item 3): auto-lookup from SRS, both tables, tag-stripped
sw_version_full = None
if srs_path and variant_info:
    st.header("3. SW 버전 (SRS)")
    vehicle = variant_info["model"].replace("PE", " PE").strip()
    phases = srs_reader.list_phases(srs_path, vehicle)
    if phases:
        phase = st.selectbox("Phase", phases)
        sw_version_full = srs_reader.find_sw_version(srs_path, vehicle, phase)
        if sw_version_full:
            st.success(f"SW Version: {sw_version_full}")
    else:
        st.warning(f"SRS에서 '{vehicle}' 차종을 못 찾음. 직접 입력하세요.")
        sw_version_full = st.text_input("SW Version (수동)")

# -- 4. People / fixed / choices
st.header("4. 선택 항목")
# PM (concept item 2): English name dropdown -> written as "Lastname, Firstname"
roster = cfg.get("pm_roster", [])
principal = ""
if roster:
    opts = [f'{p["label"]}  ·  {p["org"]} ({p["title"]})' for p in roster]
    idx = st.selectbox("Principal (PM)", range(len(opts)), format_func=lambda i: opts[i])
    principal = roster[idx]["sstr"]
    st.caption(f"→ SSTR 기입값: {principal}")
else:
    principal = st.text_input("Principal (PM, 영문)")

c1, c2 = st.columns(2)
test_type = c1.selectbox("Test Type", cfg["choices"]["test_type"])
test_result = c2.selectbox("Test Result", cfg["choices"]["test_result"])
reason = st.selectbox("Reason for Test", cfg["choices"]["reason_for_test"])
hw_version = st.text_input("HW-Version", "1.00")
pcb_index = st.text_input("PCB-Index", "N/A")

# -- 5. B. Technical Documents (concept item 5): spec name + revision (auto, editable)
st.header("5. B. Technical Documents")
tech_docs = []
if srs_path:
    rev = srs_reader.load_latest_revision(srs_path)
    default_rev = rev["ver"] if rev else ""
    if rev:
        st.caption(f"SRS 최신 리비전 자동 조회: {rev['ver']} ({rev['released']})")
    crs_name = st.text_input("CRS / 사양서 이름", "SRS_GEN2")
    crs_rev = st.text_input("Doors Revision", default_rev)  # editable (e.g. 2.08 vs 2.09)
    if crs_name and crs_rev:
        tech_docs.append({"name": crs_name, "revision": crs_rev})

# -- 6. Jira (objective + fixVersion capture/link for B. Technical Documents)
st.header("6. Jira (선택)")
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

st.header("7. SSTR 생성")
if st.button("자동 채움 실행", type="primary"):
    if not (terra_path and variant_info):
        st.error("teRA 템플릿 + 엑셀(차종) 필수.")
    else:
        terra = terra_reader.load_terra_fields(terra_path)

        # Project Designation / Order: 3rd-gen rule (concept item 1)
        gen3 = srs_reader.resolve_gen3(variant_info["model"], cfg) \
            if hasattr(srs_reader, "resolve_gen3") else None
        if gen3:
            designation = gen3["designation"]
            order = gen3["order"]
        else:
            designation = variant_info["model"].replace("PE", " PE")
            order = ""

        sw_full = (
            (terra.get("sw_version") or "")
            + ("\n" + sw_version_full if sw_version_full else "")
        ).strip()

        data = {
            "Customer": variant_info["customer"],
            "Project Designation: (Project/Type)": designation,
            "Order": order,
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
        }
        paragraphs = {
            "Objective": objective_key,
            "Result": test_result,
        }
        out = os.path.join(tempfile.gettempdir(), "SSTR_filled.docx")
        fill_sstr(terra_path, out, data, paragraphs=paragraphs,
                  reason=reason, terra=terra, tech_docs=tech_docs)
        with open(out, "rb") as f:
            st.download_button("📥 완성본 다운로드", f, file_name="SSTR_filled.docx")
        st.success("생성 완료")
