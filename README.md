# SSTR Auto-Fill

teRA 추출본 + eepxtool 엑셀 + SRS + Jira 를 통합해 SSTR Word 문서를 자동 채움한다.

## 설치
```bash
pip install -r requirements.txt
cp .env.example .env   # .env 에 JIRA_PAT 입력
streamlit run app.py
```

## 소스별 매핑
| SSTR 필드 | 소스 |
|---|---|
| Customer / Part Number / Model | 엑셀 Variant_1 |
| SW 빌드스트링 | SRS §7.4.4 |
| Objective / fixVersion | Jira (PAT) |
| Test period / Metrics / D.Matrix | teRA (복사) |
| PM / HW / Test Type / Result | 사용자 입력 |

## 프로젝트 구조
```
SSTR-document-generator/
├── config.yaml
├── .env.example
├── requirements.txt
├── sources/
│   ├── __init__.py
│   ├── terra_reader.py
│   ├── excel_reader.py
│   ├── srs_reader.py
│   └── jira_client.py
├── filler.py
└── app.py
```

## 엑셀 주의
`eepxtool` 파일은 확장자가 `.xlsx` 여도 실제 `.xls`(OLE2) 인 경우가 있어
매직바이트로 포맷을 판별해 xlrd/openpyxl 을 자동 선택한다.
